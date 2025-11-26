VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAction 
   Caption         =   "RezDruck"
   ClientHeight    =   3615
   ClientLeft      =   390
   ClientTop       =   945
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
   Icon            =   "winrezdr.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   3615
   ScaleWidth      =   2745
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
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

Private Const DefErrModul = "WINREZK.FRM"

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
Static InResize%

If (InResize%) Then Call DefErrPop: Exit Sub

InResize% = True

If (Me.Visible) Then
    Me.WindowState = vbMinimized
End If

InResize% = False

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
    l& = GetPrivateProfileString("Rezeptkontrolle", "AvpTeilnahme", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        AvpTeilnahme% = True
    Else
        AvpTeilnahme% = False
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
Dim i%, j%
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

For i% = 0 To (Printers.Count - 1)
    h$ = Printers(i%).DeviceName
    If (h$ = RezeptDrucker$) Then
        Set Printer = Printers(i%)
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
