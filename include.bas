Attribute VB_Name = "modInclude"
Option Explicit

Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
(ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long

Declare Function MsgBox2 Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByValnXStart As Long, ByVal nYStart As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long

Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_SYNCHRONIZE = &H100000

Public Const WM_CLOSE = &H10

Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Declare Sub RoundRect Lib "gdi32" (ByVal hdc&, ByVal x1&, ByVal y1&, ByVal x2&, ByVal Y2&, ByVal X3&, ByVal Y3&)
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1


Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

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

Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Const COLOR_SCROLLBAR = 0 'The Scrollbar colour
Const COLOR_BACKGROUND = 1 'Colour of the background with no wallpaper
Const COLOR_ACTIVECAPTION = 2 'Caption of Active Window
Const COLOR_INACTIVECAPTION = 3 'Caption of Inactive window
Public Const COLOR_MENU = 4 'Menu
Const COLOR_WINDOW = 5 'Windows background
Const COLOR_WINDOWFRAME = 6 'Window frame
Const COLOR_MENUTEXT = 7 'Window Text
Const COLOR_WINDOWTEXT = 8 '3D dark shadow (Win95)
Const COLOR_CAPTIONTEXT = 9 'Text in window caption
Const COLOR_ACTIVEBORDER = 10 'Border of active window
Const COLOR_INACTIVEBORDER = 11 'Border of inactive window
Const COLOR_APPWORKSPACE = 12 'Background of MDI desktop
Const COLOR_HIGHLIGHT = 13 'Selected item background
Const COLOR_HIGHLIGHTTEXT = 14 'Selected menu item
Const COLOR_BTNFACE = 15 'Button
Const COLOR_BTNSHADOW = 16 '3D shading of button
Const COLOR_GRAYTEXT = 17 'Grey text, of zero if dithering is used.
Const COLOR_BTNTEXT = 18 'Button text
Const COLOR_INACTIVECAPTIONTEXT = 19 'Text of inactive window
Const COLOR_BTNHIGHLIGHT = 20 '3D highlight of button
Const COLOR_2NDACTIVECAPTION = 27 'Win98 only: 2nd active window color
Const COLOR_2NDINACTIVECAPTION = 28 'Win98 only: 2nd inactive window color


Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long

Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_LOCAL = &H2
Public Const PRINTER_ENUM_NETWORK = &H40
Public Const PRINTER_ENUM_REMOTE = &H10
Public Const PRINTER_ENUM_SHARED = &H20




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




Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, _
      ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, _
      ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, _
      ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Public Const SRCCOPY = &HCC0020


Declare Function GetForegroundWindow Lib "user32" () As Long

Declare Function GetActiveWindow Lib "user32.dll" () As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Public Const SM_CYMENU = 15
Public Const SM_CYCAPTION = 4
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXVSCROLL = 2





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
Public Const SPI_GETNONCLIENTMETRICS = 41




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


Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously


Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Const HORZRES = 8            '  Horizontal width in pixels
Public Const HORZSIZE = 4           '  Horizontal size in millimeters
Public Const VERTRES = 10           '  Vertical width in pixels
Public Const VERTSIZE = 6           '  Vertical size in millimeters

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long


Public Const RAS_MaxEntryName = 256

Declare Function RasEnumEntries Lib "RasApi32.DLL" Alias "RasEnumEntriesA" (ByVal Reserved As String, _
   ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long



Public Const DIAL_FORCE_ONLINE = 1
Public Const DIAL_FORCE_UNATTENDED = 2

Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConiID _
    As String, ByVal dwFlags As Long, ByRef hCon As Long, ByVal dwReserved As Long) As Long
Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
Declare Function RasEnumConnectionsA Lib "RasApi32.DLL" (lpRasConn As Any, lbcp As Long, _
    lbcConnections As Long) As Long

Type RAS_ENTRYNAME
    dwSize As Long
    szEntryName(257) As Byte ' RAS_MaxEntryName + 1
    wPad1 As Integer
End Type

Type RAS_STATUS
    dwSize As Long
    hRasConn As Long
    szEntryName(256) As Byte
    szDeviceType(16) As Byte
    szDeviceName(128) As Byte
End Type



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



Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Public Const MAX_FILENAME_LEN = 260




Type MenuStruct
    Name As String * 50
    kurz As String * 2
    len As Integer
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
    aep As String * 8
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
    NNAEP As String * 8
    nnart As String * 1
    dummy As String * 40
End Type

Type TxfStruct
    Handle As Integer
    len As Long
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
    vk As Long
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

Private Const LF_FACESIZE = 32
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type







Type BedingungenStruct
    wert1 As String * 20
    op As String * 20
    wert2 As String * 20
End Type

Type KontrollenStruct
    Bedingung As BedingungenStruct
    Send As String * 1
End Type

Type ZuordnungenStruct
    Bedingung As BedingungenStruct
    lief(20) As Integer 'Byte
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

Type TaetigkeitenStruct
    Taetigkeit As String * 20
    pers(80) As Byte
End Type

Type SignaturenStruct
    Taetigkeit As String * 20
    Aktiv As Byte
End Type

Type AufschlaegeStruct
    PreisBasis As Byte
    Aufschlag As Byte
End Type

Type RundungenStruct
    Bedingung As BedingungenStruct
    Gerundet As String * 4
End Type

Type WuSortierungenStruct
    Bedingung As BedingungenStruct
End Type

Type FeiertageStruct
    Name As String * 25
    KalenderTag As String * 10
    Aktiv As String * 1
End Type

Type RowaLsStruct
    Beleg As String * 12
    LifDat As String * 9
    Ln As String * 1
    pos As Integer
End Type

Type WumsatzStruct
    lief As Byte
    bdatum As String * 6
    Wert As Double
    Rabatt As Integer
End Type

Type DruckSpalteStruct
    Titel As String * 20
    TypStr As String * 50
    BreiteX As Integer
    StartX As Long
    Ausrichtung As String * 1
    Attrib As Byte
End Type

Type VerfallWarnungStruct
    Laufzeit As Integer
    Warnung As Integer
End Type

Type SonderBelegeStruct
    pzn As String * 8
    KkBez As String * 20
    KassenId As String * 7
    Status As String * 10
    GültigBis As String * 10
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


Public Const TEXT_BESTELLSTATUS = "Artikel-Status"
Public Const TEXT_STATISTIK = "Statistik"
Public Const TEXT_SONDERANGEBOTE = "Angebote"
Public Const TEXT_ABHOLER = "Besorgerinfo"
Public Const TEXT_NACHBEARBEITUNG = "Nachbearbeitung"
Public Const TEXT_ABSAGEN = "Absagen"

Public Const TEXT_WU_BUCHEN = "Lagerstand buchen"
Public Const TEXT_WU_PREIS = "Preiskalkulation"

Public Const ZUSATZ_ARTIKEL = "Artikel"
Public Const ZUSATZ_LIEFERANTEN = "Lieferanten"

Public Const EURO_FAKTOR = 1.95583

Public Const KZ_ORIGINAL = &H80
Public Const KZ_IMPORT = &H40
Public Const KZ_BTM = &H20
Public Const KZ_KALT = &H10
Public Const KZ_ZUSTEXT = &H8
Public Const KZ_MERKZETTEL = &H4
Public Const KZ_SELBSTANGELEGT = &H2
Public Const KZ_AUSSERHANDEL = &H1

Public Const KZ_DOKUPFLICHT = &H20
Public Const KZ_NICHTAM = &H10
Public Const KZ_ISTREZPFLICHTIG = &H8
Public Const KZ_ISTINTERIM = &H4
Public Const KZ_GIBTPREISGUENSTIG = &H2
Public Const KZ_ISTPREISGUENSTIG = &H1

Public Const ABHOLER_DB = "abholer.mdb"


Public Const NO_ERROR = 0&
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

