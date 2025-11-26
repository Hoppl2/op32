VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSenden 
   Caption         =   "Bestellung senden"
   ClientHeight    =   7020
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmSenden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9855
   Begin VB.CommandButton cmdEscSend 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   7920
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   3600
      TabIndex        =   9
      Top             =   6480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Senden"
      Default         =   -1  'True
      Height          =   450
      Left            =   2160
      TabIndex        =   8
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Frame fmeAuftrag 
      Caption         =   "Auftrag"
      Height          =   4335
      Left            =   360
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboAuftrag 
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox cboAuftrag 
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkAuftrag 
         Caption         =   "Absagen &berücksichtigen"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   4335
      End
      Begin VB.CheckBox chkAuftrag 
         Caption         =   "&Rückmeldungen anzeigen"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   3480
         Width           =   4335
      End
      Begin VB.OptionButton optAuftrag 
         Caption         =   "&Warten auf Anruf"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   3655
      End
      Begin VB.OptionButton optAuftrag 
         Caption         =   "Anruf &durchführen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   3655
      End
      Begin VB.TextBox txtAuftrag 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtAuftrag 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAuftrag 
         Caption         =   "Auftrags&art"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label lblAuftrag 
         Caption         =   "Auftrags&ergänzung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Frame fmeLieferant 
      Caption         =   "Lieferant"
      Height          =   3615
      Left            =   5640
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Apotheke"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Lieferant"
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Caption         =   "Tel-Lieferant"
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   1980
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   2340
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   2700
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   3060
         Width           =   3975
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2400
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   9000
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DTREnable       =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid flxAuftrag 
      Height          =   540
      Left            =   2040
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   953
      _Version        =   65541
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin ComCtl2.Animation aniSenden 
      Height          =   1095
      Left            =   3360
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1931
      _Version        =   327681
      Center          =   -1  'True
      FullWidth       =   337
      FullHeight      =   73
   End
   Begin VB.Label lblModem 
      Caption         =   "Modem:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblModemWert 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   7695
   End
End
Attribute VB_Name = "frmSenden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "frmsenden.frm"

Private Sub cmdEsc_Click()

If (seriell% And (comSenden.PortOpen)) Then comSenden.PortOpen = False
    
Call UpdateBekartDat(Lieferant%, False)
Unload Me

End Sub

Private Sub cmdEscSend_Click()

BestSendenAbbruch% = True

End Sub

Private Sub cmdOk_Click()

Call SetFormModus(1)
Call SendGermany

End Sub

Private Sub Form_Load()
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$


Call InitBestellung

Call wpara.InitFont(Me)


fmeAuftrag.Left = wpara.LinksX%
fmeAuftrag.Top = wpara.TitelY%

txtAuftrag(0).Top = 2 * wpara.TitelY%
For i% = 1 To 1
    txtAuftrag(i%).Top = txtAuftrag(i% - 1).Top + txtAuftrag(i% - 1).Height + 90
Next i%

lblAuftrag(0).Left = wpara.LinksX%
lblAuftrag(0).Top = txtAuftrag(0).Top
For i% = 1 To 1
    lblAuftrag(i%).Left = lblAuftrag(i% - 1).Left
    lblAuftrag(i%).Top = txtAuftrag(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblAuftrag(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtAuftrag(0).Left = lblAuftrag(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtAuftrag(i%).Left = txtAuftrag(i% - 1).Left
Next i%

optAuftrag(0).Left = wpara.LinksX%
optAuftrag(0).Top = txtAuftrag(1).Top + txtAuftrag(1).Height + 300
For i% = 1 To 1
    optAuftrag(i%).Left = optAuftrag(i% - 1).Left
    optAuftrag(i%).Top = optAuftrag(i% - 1).Top + optAuftrag(i% - 1).Height + 45
Next i%

chkAuftrag(0).Left = wpara.LinksX%
chkAuftrag(0).Top = optAuftrag(1).Top + optAuftrag(1).Height + 300
For i% = 1 To 1
    chkAuftrag(i%).Left = chkAuftrag(i% - 1).Left
    chkAuftrag(i%).Top = chkAuftrag(i% - 1).Top + chkAuftrag(i% - 1).Height + 45
Next i%



With cboAuftrag(0)
    .Width = TextWidth("Zeilenwert-Auftrag (ZW)") + 300 + wpara.FrmScrollHeight
    
    .Clear
    .AddItem "Zustellung Heute (ZH)"
    .AddItem "Zustellung Morgen (ZM)"
    .AddItem "Heute kein Auftrag (KA)"
    
    .Top = txtAuftrag(0).Top
    .Left = txtAuftrag(0).Left
    .ListIndex = 0
    .Visible = True
End With

With cboAuftrag(1)
    .Width = cboAuftrag(0).Width
    
    .Clear
    .AddItem "Normalauftrag (  )"
    .AddItem "Inventurliste (IN)"
    .AddItem "Lochkarten (LK)"
    .AddItem "Rückkauf-Anfrage (RK)"
    .AddItem "SBL-Auftrag (SB)"
    .AddItem "Sonder-Auftrag (SO)"
    .AddItem "Stapel-Auftrag (ST)"
    .AddItem "Test-Auftrag (TE)"
    .AddItem "Verfalldatenliste (VD)"
    .AddItem "Vorratskauf (VR)"
    .AddItem "10er-Auftrag (ZE)"
    .AddItem "Zeit-Auftrag (ZT)"
    .AddItem "Zeilenwert-Auftrag (ZW)"
    
    .Top = txtAuftrag(1).Top
    .Left = txtAuftrag(1).Left
    .ListIndex = 0
    .Visible = True
End With


txtAuftrag(0).Visible = False
txtAuftrag(1).Visible = False



'fmeAuftrag.Width = txtAuftrag(0).Left + txtAuftrag(0).Width + 2 * wpara.LinksX%
fmeAuftrag.Width = cboAuftrag(1).Left + cboAuftrag(1).Width + 2 * wpara.LinksX%
Hoehe1% = chkAuftrag(1).Top + chkAuftrag(1).Height

''''''''''''''''

fmeLieferant.Left = fmeAuftrag.Left + fmeAuftrag.Width + 300
fmeLieferant.Top = wpara.TitelY%

lblLieferant(0).Top = 2 * wpara.TitelY%
For i% = 1 To 2
    lblLieferant(i%).Top = lblLieferant(i% - 1).Top + lblLieferant(i% - 1).Height + 60
Next i%
lblLieferant(3).Top = lblLieferant(2).Top + lblLieferant(2).Height + 180
For i% = 4 To 6
    lblLieferant(i%).Top = lblLieferant(i% - 1).Top + lblLieferant(i% - 1).Height + 15
Next i%

lblLieferant(0).Left = wpara.LinksX%
lblLieferantWert(0).Top = lblLieferant(0).Top
For i% = 1 To 6
    lblLieferant(i%).Left = lblLieferant(i% - 1).Left
    If (i < 3) Then
        lblLieferantWert(i%).Top = lblLieferant(i%).Top
    End If
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblLieferant(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblLieferantWert(0).Left = lblLieferant(0).Left + MaxWi% + 300
For i% = 1 To 2
    lblLieferantWert(i%).Left = lblLieferantWert(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblLieferantWert(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
MaxWi% = MaxWi% + lblLieferantWert(0).Left - lblLieferant(0).Left
For i% = 3 To 6
    wi% = lblLieferant(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

fmeLieferant.Width = MaxWi% + 2 * wpara.LinksX%
'fmeLieferant.Width = lblLieferantWert(0).Left + lblLieferantWert(0).Width + 2 * LinksX%
Hoehe2% = lblLieferant(6).Top + lblLieferant(6).Height


If (Hoehe1% > Hoehe2%) Then
    fmeAuftrag.Height = Hoehe1% + wpara.TitelY%
Else
    fmeAuftrag.Height = Hoehe2% + wpara.TitelY%
End If
fmeLieferant.Height = fmeAuftrag.Height


cmdOk.Top = fmeLieferant.Top + fmeLieferant.Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = fmeLieferant.Left + fmeLieferant.Width + 2 * wpara.LinksX%

cmdOk.Width = wpara.ButtonX%
cmdOk.Height = wpara.ButtonY%
cmdEsc.Width = wpara.ButtonX%
cmdEsc.Height = wpara.ButtonY%
cmdOk.Left = (ScaleWidth - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2



'''''''''''''''
lblModem.Left = wpara.LinksX%
lblModem.Top = wpara.TitelY%

lblModemWert.Left = lblModem.Left + lblModem.Width + 300
lblModemWert.Top = lblModem.Top


aniSenden.Top = cmdOk.Top - aniSenden.Height - 150
aniSenden.Left = (ScaleWidth - aniSenden.Width) / 2

flxAuftrag.Left = wpara.LinksX%
flxAuftrag.Top = lblModem.Top + lblModem.Height + 300
flxAuftrag.Width = ScaleWidth - 2 * wpara.LinksX%
flxAuftrag.Height = aniSenden.Top - flxAuftrag.Top - 150
flxAuftrag.ColWidth(0) = flxAuftrag.Width
flxAuftrag.ColAlignment(0) = flexAlignLeftCenter

cmdEscSend.Top = cmdEsc.Top
cmdEscSend.Left = (ScaleWidth - cmdEscSend.Width) / 2
cmdEscSend.Width = wpara.ButtonX%
cmdEscSend.Height = wpara.ButtonY%

''''''''''''''

Call SetFormModus(0)

End Sub

Private Sub txtAuftrag_GotFocus(Index As Integer)

With txtAuftrag(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

End Sub

Sub StartAnimation(c As Animation, Optional text$ = "Aufgabe wird bearbeitet ...")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StartAnimation")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
MousePointer = vbHourglass
aniSenden.Open "filecopy.avi"
aniSenden.Play
Refresh

Call DefErrPop
End Sub

Sub SetFormModus(modus%)

If (modus% = 0) Then
    lblModem.Visible = False
    lblModemWert.Visible = False
    flxAuftrag.Visible = False
    aniSenden.Visible = False
    cmdEscSend.Visible = False

    fmeAuftrag.Visible = True
    fmeLieferant.Visible = True
    cmdEsc.Visible = True
    cmdEsc.Cancel = True
    cmdOk.Visible = True
    cmdOk.Default = True
Else
    fmeAuftrag.Visible = False
    fmeLieferant.Visible = False
    cmdEsc.Visible = False
    cmdOk.Visible = False
    
    lblModem.Visible = True
    lblModemWert.Visible = True
    flxAuftrag.Visible = True
    flxAuftrag.Rows = 0
    aniSenden.Visible = True
    cmdEscSend.Visible = True
    cmdEscSend.Cancel = True
End If

End Sub

