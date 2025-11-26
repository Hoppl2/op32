VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAction 
   Caption         =   "WinRexX"
   ClientHeight    =   3615
   ClientLeft      =   1950
   ClientTop       =   3315
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "winrezx.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   3615
   ScaleWidth      =   2745
   Begin VB.PictureBox picTaxierungsDruck 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   1320
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboVerfügbarkeit 
      Height          =   360
      Index           =   2
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboVerfügbarkeit 
      Height          =   360
      Index           =   1
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboVerfügbarkeit 
      Height          =   360
      Index           =   0
      Left            =   0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   840
   End
   Begin VB.ListBox lstSortierung 
      Height          =   300
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Rezeptkontrolle"
Const INFO_SECTION = "Infobereich Rezeptkontrolle"


'Dim scrAuswahlAltValue%
Dim InRowColChange%

Dim Standard%

Dim ClickOk%

Private Const DefErrModul = "WINREZX.FRM"

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Load")
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
Dim i%
Dim l&
Dim h$

'If (App.PrevInstance) Then End
'If (Command$ = "") Then End

Call wpara.InitEndSub(Me)

Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI, "WinRezDr")
Call wpara.InitFont(Me)
Call HoleIniWerte

'Call WinArtDebug("vor InitProgrammTyp")
Call InitProgrammTyp

'Me.SetFocus
'DoEvents

Me.WindowState = vbMinimized

Call DefErrPop
End Sub

Sub InitProgrammTyp()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitProgrammTyp")
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
Dim i%, h$

Set ActProgram = New clsWinRezDr

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
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
    
Select Case (x \ Screen.TwipsPerPixelX)
    Case WM_LBUTTONDOWN
        If (Me.Visible = False) Then
            Me.WindowState = vbNormal
            Me.Visible = True
'            Height = picProgress.Top + picProgress.Height + wpara.TitelY + 90 + wpara.FrmScrollHeight + wpara.FrmScrollHeight
            Width = 3000
            Height = wpara.FrmCaptionHeight + wpara.FrmMenuHeight + 90
'            Call FensterImVordergrund
        End If
    Case WM_RBUTTONDOWN
    Case WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK
End Select

Call DefErrPop
End Sub

Private Sub mnuBeenden_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBeenden_Click")
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
'Close #DruckFile%

Unload Me

Call DefErrPop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_QueryUnload")
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
If (comSenden.PortOpen) Then comSenden.PortOpen = False
Call ProgrammEnde
Call DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Resize")
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
'Static InResize%
'
'If (InResize%) Then Call DefErrPop: Exit Sub
'
'InResize% = True
'
'If (Me.Visible) Then
'    Me.WindowState = vbMinimized
'End If
'
'InResize% = False

Me.Visible = False

Call ShowSysTrayIcon(frmAction, 1, frmAction.Icon, RTrim$(frmAction.Caption))

'tmrStart.Enabled = True

Call DefErrPop
End Sub

Sub frmActionUnload()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("frmActionUnload")
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

Call DefErrPop
End Sub

Sub HoleIniWerte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniWerte")
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
Dim i%, j%, k%, ind%, iVal%, erg%
Dim l&, f&, lVal&, lColor&
Dim h$, h2$, Key$, wert1$, BetrLief$, Lief2$
Dim sPzn$, sKkBez$, sKassenId$, sStatus$, sGültigBis$
    
With Me
    
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerVersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptVersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptVersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "PrivatRezeptDruckerVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    PrivatRezeptVersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDatumVersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    DatumVersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptNrVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptNrVersatzY% = Val(h$)
            
'    h2$ = Left$("Lucida Console" + Space$(20), 20)
    h2$ = Left$("Courier New" + Space$(20), 20)
    h$ = Space$(20)
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerSchriftArt", h2$, h$, 21, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptFont$ = RTrim$(h$)
    

    h$ = Space$(50)
    l& = GetPrivateProfileString(UserSection$, "RezeptDrucker", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    RezeptDrucker$ = h$
    
    h$ = Space$(50)
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerParameter", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    RezeptDruckerPara$ = h$


''' Code 128
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "Code128VersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    Code128VersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "Code128VersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    Code128VersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    h2$ = Left$("Arial,8" + Space$(20), 20)
    h$ = Space$(20)
    l& = GetPrivateProfileString(UserSection$, "Code128SchriftArt", h2$, h$, 21, INI_DATEI)
    h$ = Left$(h$, l&)
    ind% = InStr(h$, ",")
    Code128Font$ = RTrim$(Left$(h$, ind% - 1))
    Code128FontSize = Val(Mid$(h$, ind% + 1))
'''''''''''''''''''''''''''
    


    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "InstitutsKz", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    If (h$ <> "") Then h$ = Right$(Space$(7) + Trim(Left$(h$, l&)), 7)
    RezApoNr$ = h$
    OrgRezApoNr$ = RezApoNr$
    
    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezeptText", h$, h$, 51, INI_DATEI)
    RezApoDruckName$ = Trim(Left$(h$, l&))
    
    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "BtmRezeptText", h$, h$, 51, INI_DATEI)
    BtmRezDruckName$ = Trim(Left$(h$, l&))
    If (BtmRezDruckName$ = "") Then
        BtmRezDruckName$ = RezApoDruckName$
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "DSKNurRezeptnummerDrucken", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        DSKNurRezeptnummerDrucken% = True
    Else
        DSKNurRezeptnummerDrucken% = False
    End If

    h$ = "J"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezeptDetect", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    h$ = "N"
    If (h$ = "J") Then
        RezeptDetect% = True
    Else
        RezeptDetect% = False
    End If
    
    h$ = "10"
    l& = GetPrivateProfileString(INI_SECTION, "RezeptDruckPause", "10", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptDruckPause% = Val(h$)

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "BtmAlsZeile", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        BtmAlsZeile% = True
    Else
        BtmAlsZeile% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezepturDruck", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        RezepturDruck% = True
    Else
        RezepturDruck% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "AvpTeilnahme", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        AvpTeilnahme% = True
    Else
        AvpTeilnahme% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "Code128Aktiv", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        Code128Flag% = True
    Else
        Code128Flag% = False
    End If

    RezeptNrPositionAlt = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezeptNrPositionAlt", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        RezeptNrPositionAlt = True
    End If

    Pzn8Test = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "Pzn8Test", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        Pzn8Test = True
    End If

    
    j% = 0
    For i% = 1 To 10
        h$ = Space$(100)
        Key$ = "SonderBeleg" + Format(i%, "00")
        l& = GetPrivateProfileString("Rezeptkontrolle", Key$, " ", h$, 101, INI_DATEI)
        h$ = Trim$(h$)
        If (Len(h$) > 1) Then
            h$ = Left$(h$, Len(h$) - 1)
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sPzn$ = RTrim$(Left$(h$, ind% - 1))
                If (Len(sPzn) = 7) Then
                    sPzn = "0" + sPzn
                End If
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sKkBez$ = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sKassenId$ = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sStatus = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sGültigBis$ = RTrim$(Left$(h$, ind% - 1))
                With SonderBelege(j%)
                    .pzn = sPzn$
                    .KkBez = sKkBez$
                    .KassenId = sKassenId$
                    .Status = sStatus$
                    .GültigBis = sGültigBis$
                    j% = j% + 1
                End With
            End If
        End If
    Next i%
    AnzSonderBelege% = j%

    h$ = Space$(2000)
    l& = GetPrivateProfileString("Rezeptkontrolle", "Beigetreten", h$, h$, 2001, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    If (h$ <> "") Then h$ = "," + h$ + ","
    BeigetreteneVereinbarungen$ = h$

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "DatumObenPrivatRezept", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        DatumObenPrivatRezept = True
    Else
        DatumObenPrivatRezept = False
    End If

    DruckDebugAktiv = 0
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "DruckDebug", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        DruckDebugAktiv = True
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Allgemein", "PreisKz_62_70", "N", h$, 2, CurDir + "\FiveRx.ini")
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        PreisKz_62_70 = True
    Else
        PreisKz_62_70 = False
    End If

    AlleRezepturenMitHash = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "AlleRezepturenMitHash", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        AlleRezepturenMitHash = True
    End If

End With

Call DefErrPop
End Sub

Sub SpeicherIniWerte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniWerte")
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

Call DefErrPop
End Sub

Sub EndeDll()
End
End Sub

Sub ErzeugeDruckerAuswahl()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeDruckerAuswahl")
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
Dim i%, j%, k%
Dim h$, DosDrucker$(3)

Call WinArtDebug("ErzeugeDruckerAuswahl")

IstDosDrucker% = False

DosDrucker$(0) = "TM290"
DosDrucker$(1) = "TM290-II"
DosDrucker$(2) = "TM-U950"
DosDrucker$(3) = "TM5000II"

For i% = 0 To 3
    h$ = DosDrucker$(i%)
    If (h$ = RezeptDrucker$) Then
        IstDosDrucker% = True
    End If
Next i%

'For i% = 0 To (Printers.Count - 1)
'    h$ = wpara.PrinterNameOP(Printers(i%).DeviceName)
'    If (h$ = RezeptDrucker$) Then
'        Set Printer = Printers(i%)
'    End If
'Next i%
Call wpara.InstalledPrinters
For i = 0 To (wpara.PrinterCount - 1)
    h$ = wpara.PrinterName(i + 1)
    h = wpara.PrinterNameOP(h)
    If (UCase(h$) = UCase(RezeptDrucker$)) Then
        For k% = 0 To (Printers.Count - 1)
            If (UCase(wpara.PrinterName(i + 1)) = UCase(Printers(k).DeviceName)) Then
                Set Printer = Printers(k%)
                Exit For
            End If
        Next k
    End If
Next i%
    
Call WinArtDebug("ErzeugeDruckerAuswahl ENDE")

Call DefErrPop
End Sub

'Private Sub Timer1_Timer()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Timer1_Timer")
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
'
''On Error Resume Next
'
'Timer1.Enabled = False
'
'DruckDebugAktiv = True
'Call WinArtDebug(vbCrLf + "Programmstart")
'
'
'If (Dir$("fistam.dat") = "") Then ChDir "\user"
'INI_DATEI = CurDir + "\winop.ini"
'
'
'Set ast = New clsStamm
'Set taxe = New clsTaxe
'Set kiste = New clsKiste
'Set para = New clsOpPara
'Set wpara = New clsWinPara
'Set vk = New clsVerkauf
'Set RezTab = New clsVerkRtab
'Set VmPzn = New clsVmPzn
'Set VmBed = New clsVmBed
'Set VmRech = New clsVmRech
'
'UserSection$ = "Computer" + Format(Val(para.User))
'Call wpara.HoleWindowsParameter
'
'
'
'
'Call wpara.InitEndSub(Me)
'
'Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI, "WinRezDr")
'Call wpara.InitFont(Me)
'Call HoleIniWerte
'
'Call WinArtDebug("vor InitProgrammTyp")
'Call InitProgrammTyp
'
''Me.SetFocus
''DoEvents
'
'Me.WindowState = vbMinimized
'
'Call Main1
'
'Call DefErrPop
'End Sub

Private Sub tmrStart_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrStart_Timer")
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
Dim hWnd&
Static tmrCount%

tmrStart.Enabled = False
Call CheckAbfragen

tmrCount = tmrCount + 1
If (tmrCount > 60) Then
    hWnd = FindWindow(vbNullString, "FlexKasse (links)")
    If (hWnd <= 0) Then
        hWnd = FindWindow(vbNullString, "FlexKasse (rechts)")
    End If
    If (hWnd <= 0) Then
        Unload Me
        Call DefErrPop: Exit Sub
    End If
    
    tmrCount = 0
End If

tmrStart.Enabled = True

Call DefErrPop
End Sub

Sub CheckAbfragen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckAbfragen")
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
Dim i%, ind%, anz%, GhInd%, AnzPrivat%
Dim l&
Dim h$, pzn$, SQLStr$, Key$, AufrufPara$, s$, Section$

'MsgBox ("checkabf")

AufrufPara = ""
Section$ = "Computer" + Format(Val(para.User), "00")
'Section$ = "Computer" + Format(5, "00")
For i = 1 To 3
    h = Space(100)
    Key$ = "WinRezDr" + Format(i, "0")
    l& = GetPrivateProfileString(Section, Key$, h$, h$, 101, CurDir + "\BackGrnd.ini")
    h$ = Trim(Left$(h$, l&))
    If (h$ <> "") Then
        AufrufPara = h$
        Exit For
    End If
Next i

'For i = 0 To 20
'    h = Space(100)
'    key$ = "Computer" + Format(i, "00")
'    l& = GetPrivateProfileString("WinRezDr", key$, h$, h$, 101, CurDir + "\BackGrnd.ini")
'    h$ = Trim(Left$(h$, l&))
'    If (h$ <> "") Then
'        AufrufPara = h$
'        Exit For
'    End If
'
'    h = Space(100)
'    key$ = "Computer" + Format(i, "00") + "R"
'    l& = GetPrivateProfileString("WinRezDr", key$, h$, h$, 101, CurDir + "\BackGrnd.ini")
'    h$ = Trim(Left$(h$, l&))
'    If (h$ <> "") Then
'        AufrufPara = h$
'        Exit For
'    End If
'Next i

ind% = 0
If (AufrufPara$ <> "") Then
    Call WinArtDebug("vor Auswertung: " + AufrufPara)
    
    MagSpeicherIndex% = -1
    
    h$ = UCase(AufrufPara$)
    'Call MsgBox(h)
    
    HochFormatDruck = 0
    ind% = InStr(h$, "HOCH")
    If (ind% > 0) Then
        HochFormatDruck = True
        h$ = Trim(Left$(h$, ind% - 1))
    End If
    
    If (Left$(h$, 2) = "T:") Then
        h$ = UCase(Mid$(h, 3))
        If (Right$(h$, 4) <> ".MDB") Then
            h$ = h$ + ".MDB"
        End If
        VerkaufDbOk% = (Dir(h$) <> "")
        If (VerkaufDbOk) Then
'            VerkaufDB.Close
            Set VerkaufDB = OpenDatabase(h$, False, True)
            Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
            VerkaufRec.index = "Unique"
        
            AnzPrivat% = 0
            AnzRezepte% = 0
            
            RezNr$ = VerkaufRec!RezeptNr
    '        If (Left$(RezNr$, 1) = "P") Then
            If (VerkaufRec!RezeptArt = 5) Or (VerkaufRec!RezeptArt = 6) Then
                RezNr = "P" + CStr(VerkaufRec!Id)
                AnzPrivat% = AnzPrivat% + 1
                s$ = "Privat" + Str$(AnzPrivat%)
            Else
                s$ = RezNr$
            End If
            frmAction.Caption = s$
            If (ActProgram.RezeptHolen) Then
            '    Call ActProgram.RepaintBtmGebühr
            '    Call ActProgram.ShowNichtInTaxe
                Call ActProgram.DruckeRezept
            End If
            
            VerkaufDB.Close
'            Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
'            Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'            VerkaufRec.Index = "Unique"
        End If
    Else
        ind% = InStr(h$, "IK")
        If (ind% > 0) Then
            h$ = Trim(Left$(h$, ind% - 1))
        End If
        If (Right$(h$, 1) <> ",") Then
            h$ = h$ + ","
        End If
        
        AnzPrivat% = 0
        AnzRezepte% = 0
        Do
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                RezNr$ = Trim(Left$(h$, ind% - 1))
                If (Left$(RezNr$, 1) = "P") Then
                    AnzPrivat% = AnzPrivat% + 1
                    s$ = "Privat" + Str$(AnzPrivat%)
                Else
                    s$ = RezNr$
                End If
                frmAction.Caption = s$
                h$ = Mid$(h$, ind% + 1)
                If (ActProgram.RezeptHolenDB) Then
                '    Call ActProgram.RepaintBtmGebühr
                '    Call ActProgram.ShowNichtInTaxe
                    AnzRezepte% = AnzRezepte% + 1
                    Call ActProgram.DruckeRezept
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    
    l& = WritePrivateProfileString(Section, Key$, "", CurDir + "\BackGrnd.ini")
End If

Call DefErrPop
End Sub


