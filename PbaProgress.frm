VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmPbaFortschritt 
   Caption         =   "Datensätze einlesen"
   ClientHeight    =   5715
   ClientLeft      =   1680
   ClientTop       =   1560
   ClientWidth     =   7440
   ControlBox      =   0   'False
   Icon            =   "PbaProgress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin MSFlexGridLib.MSFlexGrid flxTmp 
      Height          =   1215
      Left            =   6240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   2143
      _Version        =   65541
   End
   Begin VB.PictureBox picProgress 
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
      Begin VB.Label lblFortschrittDauerWert 
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
      Begin VB.Label lblFortschrittDauerWert 
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
      Begin VB.Label lblFortschrittStatusWert 
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
         Caption         =   "Datensätze übernommen"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblFortschrittStatusWert 
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
   Begin VB.Label lblProzent 
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
Attribute VB_Name = "frmPbaFortschritt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "PBAPROGRESS.FRM"


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
PbaAnalyseAbbruch% = True
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

Me.Caption = "PBA-Analyse erstellen"
lblBestVorsStatus(1).Caption = "Bearbeitetes Datum"

Call wpara.InitFont(Me)

For i% = 0 To 1
    lblFortschrittStatusWert(i%).Caption = ""
    lblFortschrittDauerWert(i%).Caption = ""
Next i%
lblProzent.Caption = ""

fmeBestVorsStatus.Left = wpara.LinksX
fmeBestVorsStatus.Top = wpara.TitelY

lblBestVorsStatus(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsStatus(i%).Top = lblBestVorsStatus(i% - 1).Top + lblBestVorsStatus(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblFortschrittStatusWert(i%).Top = lblBestVorsStatus(i%).Top
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

lblFortschrittStatusWert(0).Left = lblBestVorsStatus(0).Left + MaxWi% + 300
For i% = 1 To 1
    lblFortschrittStatusWert(i%).Left = lblFortschrittStatusWert(i% - 1).Left
Next i%

fmeBestVorsStatus.Width = lblFortschrittStatusWert(0).Left + lblFortschrittStatusWert(0).Width + 2 * wpara.LinksX
fmeBestVorsStatus.Height = lblBestVorsStatus(1).Top + lblBestVorsStatus(1).Height + wpara.TitelY

fmeBestVorsDauer.Left = wpara.LinksX
fmeBestVorsDauer.Top = fmeBestVorsStatus.Top + fmeBestVorsStatus.Height + 300

lblBestVorsDauer(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsDauer(i%).Top = lblBestVorsDauer(i% - 1).Top + lblBestVorsDauer(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblFortschrittDauerWert(i%).Top = lblBestVorsDauer(i%).Top
Next i%

lblBestVorsDauer(0).Left = wpara.LinksX
For i% = 1 To 1
    lblBestVorsDauer(i%).Left = lblBestVorsDauer(i% - 1).Left
Next i%

lblFortschrittDauerWert(0).Left = lblFortschrittStatusWert(0).Left
For i% = 1 To 1
    lblFortschrittDauerWert(i%).Left = lblFortschrittDauerWert(i% - 1).Left
Next i%

fmeBestVorsDauer.Width = lblFortschrittDauerWert(0).Left + lblFortschrittDauerWert(0).Width + 2 * wpara.LinksX
fmeBestVorsDauer.Height = lblBestVorsDauer(1).Top + lblBestVorsDauer(1).Height + wpara.TitelY


prgBestVors.Left = wpara.LinksX
prgBestVors.Top = fmeBestVorsDauer.Top + fmeBestVorsDauer.Height + 300
prgBestVors.Width = fmeBestVorsDauer.Width

lblProzent.Left = prgBestVors.Left + (prgBestVors.Width - lblProzent.Width) / 2
lblProzent.Top = prgBestVors.Top + prgBestVors.Height + 150

picProgress.Left = wpara.LinksX
picProgress.Top = fmeBestVorsDauer.Top + fmeBestVorsDauer.Height + 300
picProgress.Width = fmeBestVorsDauer.Width
picProgress.Height = picProgress.TextHeight("99 %") + 120


'cmdEsc.Top = lblBestVorsProzent.Top + lblBestVorsProzent.Height + 150
cmdEsc.Top = picProgress.Top + picProgress.Height + 210

Me.Width = fmeBestVorsDauer.Width + 2 * wpara.LinksX + 120

cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEsc.Left = (Me.ScaleWidth - cmdEsc.Width) / 2

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

PbaAnalyseAbbruch% = False

Call DefErrPop
End Sub

