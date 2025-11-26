Attribute VB_Name = "modInclude"
Option Explicit

Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Sub RoundRect Lib "gdi32" (ByVal hdc&, ByVal x1&, ByVal y1&, ByVal X2&, ByVal Y2&, ByVal x3&, ByVal y3&)
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
      
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbcopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlCopyMemory" (hpvDest As Any, hpvSource As Any, ByVal cbcopy As Long)

'Declare Function IsdnSer% Lib "seriell.lib" Alias "_seriell" (ByVal fkt As Integer, anr As Integer, sap As Any)


Public Const DRIVE_CDROM = 5

Declare Sub DxToIEEEd Lib "mbfiee32.dll" (mbf As Double)
Declare Sub DxToIEEEs Lib "mbfiee32.dll" (mbf As Single)
Declare Sub DxToMBFd Lib "mbfiee32.dll" (ieee As Double)
Declare Sub DxToMBFs Lib "mbfiee32.dll" (ieee As Single)




Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4


'Windows-Message
Public Const WM_MOUSEMOVE = &H200

'Mausklicks
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down




Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hrgn As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
      ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
      ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, _
      ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long



Declare Function GetForegroundWindow Lib "user32" () As Long


Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CYMENU = 15
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33





Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'Declare constants used by GetWindow.
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5



Public Const SPI_GETWORKAREA = 48



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





Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Type MenuStruct
    Name As String * 50
    kurz As String * 2
    Len As Integer
End Type

'Type AusgabeStruct
'    Name As String * 100 '80
'    DSnummer As Long
'    Verweis As Long
'    LagerKz As Byte
'    pzn As Long
'    Bookmark As Variant
'    key As String * 10
'    ZusatzTextKz As Byte
'    BesorgtKz As Byte
'End Type

Type AusgabeStruct1
    pzn As String * 7
    txt As String * 35
    lief As String * 1
    bm As Integer
    asatz As Integer
    ssatz As Integer
    best As String * 1
    nm As Integer
    AEP As String * 8
    abl As String * 1
    wg As String * 1
    AVP As String * 8
    km As Integer
    absage As String * 1
    angebot As String * 1
    auto As String * 1
    alt As String * 1
    zr As String * 2
    besorger As String * 1
    NNAep As String * 8
    nnart As String * 1
    dummy As String * 40
End Type

Type TxfStruct
    Handle As Integer
    Len As Long
    Size As Long
End Type

Type SuchKriterien
    Level As Integer
    OpModus As Integer
    WasSuch As String * 80
End Type

Type LauerGhStruct
    LauerName As String * 40
    LauerNr As Integer
    OpName As String * 6
    OpNr As Integer
End Type

Type LauerTaxeStruct
    pzn As Long
    Sort As Long
    M2 As Long
    NachfolgerPZN As Long
    VorgaengerPZN As Long
    OriginalPZN As Long
    HNr As Long
    NameOderSiehe As String * 48
    Hpreis As Long
    KVAPreis As Long
    EK As Long
    VK As Long
    VKalt As Long
    GDatum As String * 2
    AlteWarengruppe As Byte
    FESTBETRAG As Long
    StoffNr As Integer
    StaerkeNr As Integer
    DFG As Integer
    GHAngebot(3) As Integer
    Indikation As String * 5
    HerstellerKB As String * 5
    menge As String * 7
    MeEinh As String * 2
    Stueckelung As String * 1
    Dar As String * 3
    BTMInhaltsstoff As String * 2
    Zuzahlung As String * 7
    Warengruppe As String * 7
    LetzteAenderung As Long
    Einfuehrung As Long
    LetztePreisaenderung As Long
    ATCCode As String * 7
    BTMMenge As String * 1
    KennZ   As String * 12
End Type

Type LauerSieheStruct
    pzn As Long
    Sort As Long
    M2 As Long
    NachfolgerPZN As Long
    VorgaengerPZN As Long
    OriginalPZN As Long
    HNr As Long
    NameOderSiehe As String * 77
    dummy As String * 83
End Type

Type ProgrammStruct
    Name As String * 40
    ProgrammChar As String * 1
    Hotkey As String * 1
End Type


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallBackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type







Type BedingungenStruct
    wert1 As String * 20
    op As String * 20
    wert2 As String * 20
End Type

Type KontrollenStruct
    bedingung As BedingungenStruct
    Send As String * 1
End Type

Type ZuordnungenStruct
    bedingung As BedingungenStruct
    lief(20) As Byte
End Type

Type RufzeitenStruct
    Lieferant As Byte
    WoTag(6) As Byte
    RufZeit As Integer
    LieferZeit As Integer
    AuftragsErg As String * 2
    AuftragsArt As String * 2
    Aktiv As String * 1
    LetztSend As Long
    Gewarnt As String * 1
End Type




Public Const MENU_F2 = 0
Public Const MENU_F3 = 1
Public Const MENU_F4 = 2
Public Const MENU_F5 = 3
Public Const MENU_F6 = 4
Public Const MENU_F7 = 5
Public Const MENU_F8 = 6
Public Const MENU_F9 = 7

Public Const MENU_SF2 = 9
Public Const MENU_SF3 = 10
Public Const MENU_SF4 = 11
Public Const MENU_SF5 = 12
Public Const MENU_SF6 = 13
Public Const MENU_SF7 = 14
Public Const MENU_SF8 = 15
Public Const MENU_SF9 = 16



Public Const DLG_BESTELL_STATUS = 0
Public Const DLG_MATCHCODE = 1
Public Const DLG_LIEFERANTEN_WAHL = 2
Public Const DLG_SENDEN = 3
Public Const DLG_BESTELL_VORSCHLAG = 4
Public Const DLG_ABOUT = 5
Public Const DLG_OPTIONEN = 6
Public Const DLG_ABHOLER = 7


Public Const TEXT_BESTELLSTATUS = "Bestellstatus"
Public Const TEXT_STATISTIK = "Statistik"
Public Const TEXT_SONDERANGEBOTE = "Sonderangebote"
Public Const TEXT_ABHOLER = "Besorgerinfo"
Public Const TEXT_NACHBEARBEITUNG = "Nachbearbeitung"
Public Const TEXT_ABSAGEN = "Absagen"

Public Const ZUSATZ_ARTIKEL = "Artikel"
Public Const ZUSATZ_LIEFERANTEN = "Lieferanten"


