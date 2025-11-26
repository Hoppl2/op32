VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSenden 
   Caption         =   "Bestellung senden"
   ClientHeight    =   7020
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "Senden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9855
   Begin VB.Timer tmrAutomatik 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   5640
   End
   Begin VB.Timer tmrSenden 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   5520
   End
   Begin VB.Frame fmeStatus 
      Caption         =   "Sendestatus"
      Height          =   1695
      Left            =   5640
      TabIndex        =   25
      Top             =   4800
      Width           =   3495
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'Kein
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   615
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid flxAuftrag 
         Height          =   540
         Left            =   840
         TabIndex        =   28
         Top             =   960
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
         Left            =   1650
         TabIndex        =   27
         Top             =   360
         Width           =   7695
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
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdEscSend 
      Caption         =   "ESC"
      Height          =   450
      Left            =   7920
      TabIndex        =   24
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
      TabIndex        =   21
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Frame fmeLieferant 
      Caption         =   "Lieferant"
      Height          =   3615
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Apotheke"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Lieferant"
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Caption         =   "Tel-Lieferant"
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   1980
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   2340
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   2700
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
   Begin ComctlLib.ImageList imgSenden 
      Left            =   3600
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Senden.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Senden.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Senden.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Senden.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Senden.frx":0F72
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSenden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TimerIndex%(5)
Dim TimerStatus%


Private Const DefErrModul = "SENDEN.FRM"

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEsc_Click")
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
If (seriell% And (comSenden.PortOpen)) Then comSenden.PortOpen = False
    
If (RueckKaufSendung%) Then
    Call UpdateRueckKaufDat(Lieferant%, False)
Else
    Call UpdateBekartDat(Lieferant%, False)
End If

Unload Me

Call DefErrPop
End Sub

Private Sub cmdEscSend_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEscSend_Click")
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
BestSendenAbbruch% = True

Call DefErrPop
End Sub

Private Sub cmdOk_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOk_Click")
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

If (Wbestk2ManuellVorbereitung%) Then
    Call SpeicherManuellSendung
    Unload Me
Else
    Call SetFormModus(1)
    Call SendGermany
End If

Call DefErrPop
End Sub

Private Sub comSenden_OnComm()
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

CommEvent% = comSenden.CommEvent

Call DefErrPop
End Sub

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$


Call InitBestellung


TimerIndex%(0) = 1
TimerIndex%(1) = 2
TimerIndex%(2) = 3
TimerIndex%(3) = 4
TimerIndex%(4) = 3
TimerIndex%(5) = 2

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
    .Width = TextWidth("Normalauftrag außer Dekade (NA)") + 300 + wpara.FrmScrollHeight
    
    .Clear
    If (para.Land = "A") Then
        .AddItem "Normalfall (  )"
        .AddItem "Kein Auftrag (KA)"
        .AddItem "Rückruf erwünscht (RR)"
        .AddItem "Später (SP)"
        .AddItem "Kein Auftrag, Rückruf (KR)"
    Else
        .AddItem "Zustellung Heute (ZH)"
        .AddItem "Zustellung Morgen (ZM)"
        .AddItem "Heute kein Auftrag (KA)"
        .AddItem "Rückruf erbeten (RR)"
    End If
    
    .Top = txtAuftrag(0).Top
    .Left = txtAuftrag(0).Left
    .ListIndex = 0
    .Visible = True
End With

With cboAuftrag(1)
    .Width = cboAuftrag(0).Width
    
    .Clear
    If (para.Land = "A") Then
        .AddItem "Normalauftrag in Dekade (N0)"
        .AddItem "Normalauftrag außer Dekade (NA)"
        .AddItem "Testauftrag in Dekade (TE)"
    Else
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
    End If
    
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
fmeStatus.Top = cmdEsc.Top
fmeStatus.Left = fmeAuftrag.Left
fmeStatus.Width = fmeLieferant.Left + fmeLieferant.Width - fmeStatus.Left

lblModem.Left = wpara.LinksX%
lblModem.Top = 2 * wpara.TitelY%

lblModemWert.Left = lblModem.Left + lblModem.Width + 300
lblModemWert.Top = lblModem.Top


With picStatus
    .Left = wpara.LinksX%
    .Top = lblModem.Top + lblModem.Height + 300
    .Width = 600
    .Height = 600
    .Picture = imgSenden.ListImages(TimerIndex%(0)).ExtractIcon
End With

With flxAuftrag
    .Left = picStatus.Left + picStatus.Width + 300
    .Top = picStatus.Top
    .Width = fmeStatus.Width - .Left - wpara.LinksX%
    .Rows = 1
    .Height = .RowHeight(0) * 5 + 90
    .ColWidth(0) = flxAuftrag.Width
    .ColAlignment(0) = flexAlignLeftCenter
    fmeStatus.Height = .Top + .Height + wpara.TitelY%
End With


cmdEscSend.Top = fmeStatus.Top + fmeStatus.Height + 150
cmdEscSend.Left = (ScaleWidth - cmdEscSend.Width) / 2
cmdEscSend.Width = wpara.ButtonX%
cmdEscSend.Height = wpara.ButtonY%

''''''''''''''

If (AutomaticSend%) Then
    h$ = "(" + Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsErg) + Space$(2), 2) + ")"
    With cboAuftrag(0)
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            If (InStr(.text, h$) > 0) Then
                Exit For
            End If
        Next i%
    End With
    
    h$ = "(" + Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsArt) + Space$(2), 2) + ")"
    With cboAuftrag(1)
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            If (InStr(.text, h$) > 0) Then
                Exit For
            End If
        Next i%
    End With
    
    If (AutomatikAktivSenden%) Then
        optAuftrag(0).Value = True
    Else
        optAuftrag(1).Value = True
    End If
    
    chkAuftrag(0).Value = 0
    chkAuftrag(1).Value = 1
    
    AutomatikFertig% = False
    AutomatikFehler$ = ""
End If


Call SetFormModus(0)

If (AutomaticSend%) Then
    tmrAutomatik.Enabled = True
ElseIf (RueckKaufSendung%) Then
    cboAuftrag(1).ListIndex = 3
    chkAuftrag(1).Value = 0
    chkAuftrag(1).Enabled = False
End If

Call DefErrPop
End Sub

Private Sub tmrAutomatik_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrAutomatik_Timer")
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
tmrAutomatik.Enabled = False
cmdOk.Value = True
Call DefErrPop
End Sub

Private Sub tmrSenden_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrSenden_Timer")
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
picStatus.Picture = imgSenden.ListImages(TimerIndex%(TimerStatus%)).ExtractIcon
TimerStatus% = TimerStatus% + 1
If (TimerStatus% > 5) Then
    TimerStatus% = 0
End If
Call DefErrPop
End Sub

Private Sub txtAuftrag_GotFocus(Index As Integer)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAuftrag_GotFocus")
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
With txtAuftrag(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

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
If (UnloadMode <> 1) Then Cancel = 1
Call DefErrPop
End Sub

Sub SetFormModus(modus%)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetFormModus")
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

If (modus% = 0) Then
    Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight
    Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
    Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
    fmeStatus.Visible = False
    tmrSenden.Enabled = False
'    lblModem.Visible = False
'    lblModemWert.Visible = False
'    flxAuftrag.Visible = False
'    aniSenden.Visible = False
    cmdEscSend.Visible = False

    fmeAuftrag.Visible = True
    fmeAuftrag.Enabled = True
    For i% = 0 To 1
        lblAuftrag(i%).Enabled = True
        cboAuftrag(i%).Enabled = True
        chkAuftrag(i%).Enabled = True
        optAuftrag(i%).Enabled = True
    Next i%
    
    fmeLieferant.Visible = True
    cmdEsc.Visible = True
    cmdEsc.Cancel = True
    cmdOk.Visible = True
    cmdOk.Default = True
Else
    Me.Height = cmdEscSend.Top + cmdEscSend.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight
    Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
    Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
    
    fmeAuftrag.Enabled = False
    For i% = 0 To 1
        lblAuftrag(i%).Enabled = False
        cboAuftrag(i%).Enabled = False
        chkAuftrag(i%).Enabled = False
        optAuftrag(i%).Enabled = False
    Next i%
'    fmeAuftrag.Visible = False
'    fmeLieferant.Visible = False
    cmdEsc.Visible = False
    cmdOk.Visible = False
    
    fmeStatus.Visible = True
    tmrSenden.Enabled = True
    TimerStatus% = 0
    flxAuftrag.Rows = 0
'    lblModem.Visible = True
'    lblModemWert.Visible = True
'    flxAuftrag.Visible = True
'    flxAuftrag.Rows = 0
'    aniSenden.Visible = True
    cmdEscSend.Visible = True
'    cmdEscSend.Cancel = True
'    Call StartAnimation(aniSenden)
End If

Call DefErrPop
End Sub

