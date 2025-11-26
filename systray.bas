Attribute VB_Name = "modSysTray"
Option Explicit

Const title = "TicTacTool"

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Const cLeftClick = 1
Private Const cRightClick = 2
Private Const cLeftDblClick = &H101
Private Const cRightDblClick = &H102

'Variable zum Übermiteln der Shell_NotifyIcon-Info
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Shell_NotifyIcon-Kommandos
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'Shell_NotifyIcon-Flags
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

''Windows-Message
'Private Const WM_MOUSEMOVE = &H200

''Mausklicks
'Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
'Private Const WM_LBUTTONDOWN = &H201     'Button down
'Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
'Private Const WM_RBUTTONDOWN = &H204     'Button down

'Shell_NotifyIcon-Funktion
Private Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Hilfsvariablen um den SingleKlick auswerten zu können
Dim NiKeyEventDblClickFlag As Integer
Dim NiKeyEventButtonID As Integer

Dim CurrNid As Long
Dim CurrIcon As Integer

'Public Sub ShowPopUpRight()
'
'Dim i As Integer
'Dim P, Q As Printer
'
'i = 0
'Set Q = Printer
'    For Each P In Printers
'        i = i + 1
'        PopUpRight(60 + i).Caption = P.DeviceName
'        PopUpRight(60 + i).Visible = True
'        PopUpRight(60 + i).Checked = False
'        If P.DeviceName = Q.DeviceName Then
'            PopUpRight(60 + i).Checked = True
'        End If
'    Next
'    PopupMenu PopUpMenuRight, , , , PopUpRight(95)
'End Sub
'
'Private Sub PopUpRight_Click(Index As Integer)
'Dim P As Printer
'
'    Select Case Index
'        Case 61 To 68
'            For Each P In Printers
'                If P.DeviceName = PopUpRight(Index).Caption Then
'                    Set Printer = P
'                End If
'            Next
'        Case 80
'            DoWinExit EWX_SHUTDOWN
'        Case 81
'            DoWinExit EWX_REBOOT
'        Case 82
'            DoWinExit EWX_LOGOFF
'        Case 83
'            PopUpRight(83).Checked = Not PopUpRight(83).Checked
'            CurrIcon = CurrIcon Xor 1
'            ChangeSysTrayIcon Me, CurrNid, Picture1(CurrIcon).Picture, title
'            ShowPopUpRight
'        Case 93
'            End
'        Case 91
'        Case 95 'abbruch
'            'macht nichts, verschwindet von selbst
'    End Select
'
'End Sub
'
'Public Sub DoWinExit(i As Long)
'Dim j As Long
'    j = i
'    If PopUpRight(83).Checked Then j = j Or EWX_FORCE
'    j = ExitWindowsEx(j, 0)
'End Sub
'
'Private Sub Form_Load()
'Dim s As String
'    s = Trim$(UCase$(Command$))
'    If s <> "" Then
'        Select Case s
'            Case "/ENDE"
'                DoWinExit EWX_SHUTDOWN
'            Case "/NEUSTART"
'                DoWinExit EWX_REBOOT
'            Case "/NEUANMELDUNG"
'                DoWinExit EWX_LOGOFF
'            Case "/ENDE!"
'                DoWinExit EWX_SHUTDOWN Or EWX_FORCE
'            Case "/NEUSTART!"
'                DoWinExit EWX_REBOOT Or EWX_FORCE
'            Case "/NEUANMELDUNG!"
'                DoWinExit EWX_LOGOFF Or EWX_FORCE
'            Case Else
'                MsgBox "'" + s + "' ist ein unbekannter Startparameter.", 16, title
'                'ShowInfo
'        End Select
'        End
'    End If
'
'    If App.PrevInstance Then
'        MsgBox title & " wird bereits ausgeführt.", 64, title
'        End
'    End If
'
'    CurrIcon = 0
'    CurrNid = 1
'    ShowSysTrayIcon Me, CurrNid, Picture1(CurrIcon).Picture, title
'End Sub
'
'Private Sub Form_Terminate()
'Dim i As Long
'    For i = 1 To CurrNid
'        KillSysTrayIcon Me, i
'    Next
'End Sub
'
'Private Sub Command1_Click()
'   CurrIcon = 0
'   CurrNid = CurrNid + 1
'   ShowSysTrayIcon Me, CurrNid, Picture1(CurrIcon).Picture, "neu"
'End Sub
'
'Private Sub Command2_Click()
'    If CurrNid > 0 Then
'        CurrIcon = (CurrIcon + 1) And 1
'        ChangeSysTrayIcon Me, CurrNid, Picture1(CurrIcon).Picture, "Geändert"
'    End If
'End Sub
'
'Private Sub Command3_Click()
'    If CurrNid > 0 Then
'        KillSysTrayIcon Me, CurrNid
'        CurrNid = CurrNid - 1
'    End If
'End Sub
'
''Kommt innerhalb einer Sekunde kein DBLCLICK-Event, dann ist es ein SINGLECLICK
'Private Sub DblClkDelay_Timer()
'    DblClkDelay.Enabled = False
'    NiKeyEvent NiKeyEventButtonID Or NiKeyEventDblClickFlag
'End Sub
'
''Eventverteiler -> hiermit Anwenderfunktionen aufrufen lassen
'Sub NiKeyEvent(iKeyEvent As Integer)
'    Select Case iKeyEvent
'        Case cLeftClick
'        Case cRightClick
'            ShowPopUpRight
'        Case cLeftDblClick
'        Case cRightDblClick
'            End
'    End Select
'End Sub
'
''die Möchtegern-Call-back-Funktion
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    Select Case x \ Screen.TwipsPerPixelX
'         Case WM_LBUTTONDOWN
'            NiKeyEventDblClickFlag = 0
'            NiKeyEventButtonID = 1
'            DblClkDelay.Enabled = False
'            DblClkDelay.Enabled = True
'         Case WM_RBUTTONDOWN
'            NiKeyEventDblClickFlag = 0
'            NiKeyEventButtonID = 2
'            DblClkDelay.Enabled = False
'            DblClkDelay.Enabled = True
'         Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK
'            NiKeyEventDblClickFlag = &H100
'      End Select
'End Sub

'Funktion zum Hinzufügen eines Icons
Private Const DefErrModul = "SYSTRAY.BAS"

Sub ShowSysTrayIcon(niForm As Form, niID As Long, niIcon As Long, niText As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowSysTrayIcon")
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
NotifyIcon NIM_ADD, niForm, niID, niIcon, niText
Call DefErrPop
End Sub

''Funktion zum Ändern eines Icons
'Public Sub ChangeSysTrayIcon(niForm As Form, niID As Long, niIcon As Long, niText As String)
'    NotifyIcon NIM_MODIFY, niForm, niID, niIcon, niText
'End Sub

'Funktion zum Löschen eines Icons
Public Sub KillSysTrayIcon(niForm As Form, niID As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KillSysTrayIcon")
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
    NotifyIcon NIM_DELETE, niForm, niID, vbNull, ""
Call DefErrPop
End Sub

'NotifyIcon-API-Aufruf
Public Sub NotifyIcon(niCmd As Long, niForm As Form, niID As Long, niIcon As Long, niText As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NotifyIcon")
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
Dim niData As NOTIFYICONDATA
Dim rc As Integer
    niData.cbSize = Len(niData)
    niData.hWnd = niForm.hWnd
    niData.uId = niID
    niData.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    niData.uCallBackMessage = WM_MOUSEMOVE
    niData.hIcon = niIcon
    niData.szTip = Left$(niText, 63) & vbNullChar
    rc = Shell_NotifyIconA(niCmd, niData)
Call DefErrPop
End Sub


