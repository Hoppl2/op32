VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmBestVors 
   Caption         =   "Bestellvorschlag"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin VB.PictureBox picBestVorsProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Stift maskieren
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Timer tmrBestVors 
      Interval        =   500
      Left            =   3960
      Top             =   4680
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Abbruch"
      Height          =   500
      Left            =   1560
      TabIndex        =   10
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Frame fmeBestVorsDauer 
      Caption         =   "Einspieldauer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   8400
      Begin VB.Label lblBestVorsDauer 
         Alignment       =   2  'Zentriert
         Caption         =   "Bisher"
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
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblBestVorsDauerWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblBestVorsDauer 
         Alignment       =   2  'Zentriert
         Caption         =   "Rest"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblBestVorsDauerWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   6120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame fmeBestVorsStatus 
      Caption         =   "Einspielstatus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5400
      Begin VB.Label lblBestVorsStatusWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblBestVorsStatus 
         Alignment       =   2  'Zentriert
         Caption         =   "Anzahl Bestellt"
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
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblBestVorsStatusWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblBestVorsStatus 
         Alignment       =   2  'Zentriert
         Caption         =   "Anzahl Datensätze"
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
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
   End
   Begin ComctlLib.ProgressBar prgBestVors 
      Height          =   255
      Left            =   -120
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblBestVorsProzent 
      Alignment       =   2  'Zentriert
      Caption         =   "999%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmBestVors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsc_Click()

BestvorsAbbruch% = True
End Sub

Private Sub Form_Load()
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wpara.InitFont(Me)

For i% = 0 To 1
    lblBestVorsStatusWert(i%).Caption = ""
    lblBestVorsDauerWert(i%).Caption = ""
Next i%
lblBestVorsProzent.Caption = ""

fmeBestVorsStatus.Left = wpara.LinksX
fmeBestVorsStatus.Top = wpara.TitelY

lblBestVorsStatus(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsStatus(i%).Top = lblBestVorsStatus(i% - 1).Top + lblBestVorsStatus(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblBestVorsStatusWert(i%).Top = lblBestVorsStatus(i%).Top
Next i%

lblBestVorsStatus(0).Left = wpara.LinksX
For i% = 1 To 1
    lblBestVorsStatus(i%).Left = lblBestVorsStatus(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblBestVorsStatus(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblBestVorsStatusWert(0).Left = lblBestVorsStatus(0).Left + MaxWi% + 300
For i% = 1 To 1
    lblBestVorsStatusWert(i%).Left = lblBestVorsStatusWert(i% - 1).Left
Next i%

fmeBestVorsStatus.Width = lblBestVorsStatusWert(0).Left + lblBestVorsStatusWert(0).Width + 2 * wpara.LinksX
fmeBestVorsStatus.Height = lblBestVorsStatus(1).Top + lblBestVorsStatus(1).Height + wpara.TitelY



fmeBestVorsDauer.Left = wpara.LinksX
fmeBestVorsDauer.Top = fmeBestVorsStatus.Top + fmeBestVorsStatus.Height + 300

lblBestVorsDauer(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsDauer(i%).Top = lblBestVorsDauer(i% - 1).Top + lblBestVorsDauer(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblBestVorsDauerWert(i%).Top = lblBestVorsDauer(i%).Top
Next i%

lblBestVorsDauer(0).Left = wpara.LinksX
For i% = 1 To 1
    lblBestVorsDauer(i%).Left = lblBestVorsDauer(i% - 1).Left
Next i%

lblBestVorsDauerWert(0).Left = lblBestVorsStatusWert(0).Left
For i% = 1 To 1
    lblBestVorsDauerWert(i%).Left = lblBestVorsDauerWert(i% - 1).Left
Next i%

fmeBestVorsDauer.Width = lblBestVorsDauerWert(0).Left + lblBestVorsDauerWert(0).Width + 2 * wpara.LinksX
fmeBestVorsDauer.Height = lblBestVorsDauer(1).Top + lblBestVorsDauer(1).Height + wpara.TitelY


prgBestVors.Left = wpara.LinksX
prgBestVors.Top = fmeBestVorsDauer.Top + fmeBestVorsDauer.Height + 300
prgBestVors.Width = fmeBestVorsDauer.Width

lblBestVorsProzent.Left = prgBestVors.Left + (prgBestVors.Width - lblBestVorsProzent.Width) / 2
lblBestVorsProzent.Top = prgBestVors.Top + prgBestVors.Height + 150

picBestVorsProgress.Left = wpara.LinksX
picBestVorsProgress.Top = fmeBestVorsDauer.Top + fmeBestVorsDauer.Height + 300
picBestVorsProgress.Width = fmeBestVorsDauer.Width
picBestVorsProgress.Height = picBestVorsProgress.TextHeight("99 %") + 120


'cmdEsc.Top = lblBestVorsProzent.Top + lblBestVorsProzent.Height + 150
cmdEsc.Top = picBestVorsProgress.Top + picBestVorsProgress.Height + 210

Me.Width = fmeBestVorsDauer.Width + 2 * wpara.LinksX + 120

cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEsc.Left = (Me.ScaleWidth - cmdEsc.Width) / 2

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.frmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

End Sub

Private Sub tmrBestVors_Timer()

tmrBestVors.Enabled = False
Call BestellVorschlag(False)
Unload Me

End Sub
