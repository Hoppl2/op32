VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAction 
   BorderStyle     =   0  'Kein
   Caption         =   "Wbestk2"
   ClientHeight    =   7890
   ClientLeft      =   -1125
   ClientTop       =   1170
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wbestk2.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   7890
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstDirektSortierung 
      Height          =   300
      Left            =   9120
      Sorted          =   -1  'True
      TabIndex        =   51
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picSchwellwertAction 
      Height          =   975
      Left            =   9480
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   47
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      Begin MSFlexGridLib.MSFlexGrid flxSchwellwertAction 
         Height          =   615
         Left            =   240
         TabIndex        =   48
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         _Version        =   393216
         ScrollBars      =   2
      End
   End
   Begin VB.PictureBox picRufzeiten 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   9255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox txtDDEServer 
         Height          =   495
         Left            =   2640
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid flxRufzeiten 
         Height          =   3960
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   6985
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   16514774
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   3960
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   6985
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   16514774
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Timer tmrAction 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   720
         Top             =   1080
      End
   End
   Begin VB.PictureBox picHinweis 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   5640
      Left            =   0
      ScaleHeight     =   5640
      ScaleWidth      =   10695
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      Begin VB.Timer tmrHinweis 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4920
         Top             =   720
      End
      Begin VB.CommandButton cmdHinweisWeg 
         Caption         =   "Streichen"
         Height          =   450
         Left            =   2400
         TabIndex        =   37
         Top             =   3000
         Width           =   1200
      End
      Begin VB.CommandButton cmdHinweisOk 
         Caption         =   "OK"
         Height          =   450
         Left            =   480
         TabIndex        =   36
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Frame fmeHinweis 
         Caption         =   "Rufzeiteneintrag"
         Height          =   2295
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   4575
         Begin VB.PictureBox picZeit 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            DrawMode        =   10  'Stift maskieren
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Ausgefüllt
            ForeColor       =   &H8000000D&
            Height          =   495
            Left            =   720
            ScaleHeight     =   435
            ScaleWidth      =   2940
            TabIndex        =   31
            Top             =   915
            Width           =   3000
         End
         Begin VB.Label lblZeitWert 
            Caption         =   "999:99:99"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   35
            Top             =   1635
            Width           =   1095
         End
         Begin VB.Label lblZeitWert 
            Caption         =   "999:99:99"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   34
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblZeit 
            Caption         =   "Aktuelle Uhrzeit: "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   1635
            Width           =   2295
         End
         Begin VB.Label lblZeit 
            Caption         =   "Eingetragene Rufzeit: "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   2295
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxSortierung 
      Height          =   975
      Left            =   9600
      TabIndex        =   46
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.Timer tmrOptimal 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   1080
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   450
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Index           =   0
      Left            =   3360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox lstSortierung 
      Height          =   300
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   1080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picSendGlobal 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   5640
      Left            =   0
      ScaleHeight     =   5640
      ScaleWidth      =   10695
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      Begin VB.Frame fmeDirektBezug 
         Caption         =   "Übertragungsart"
         Height          =   1455
         Left            =   5760
         TabIndex        =   49
         Top             =   3960
         Width           =   4575
         Begin VB.ListBox lstDirektBezug 
            Height          =   540
            Left            =   1080
            TabIndex        =   50
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame fmeLieferant 
         Caption         =   "Lieferant"
         Height          =   3615
         Left            =   5520
         TabIndex        =   18
         Top             =   0
         Width           =   5055
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
            TabIndex        =   28
            Top             =   1440
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
            TabIndex        =   27
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
            Index           =   0
            Left            =   4440
            TabIndex        =   26
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblLieferant 
            Height          =   300
            Index           =   6
            Left            =   360
            TabIndex        =   25
            Top             =   3060
            Width           =   3975
         End
         Begin VB.Label lblLieferant 
            Height          =   300
            Index           =   5
            Left            =   360
            TabIndex        =   24
            Top             =   2700
            Width           =   3975
         End
         Begin VB.Label lblLieferant 
            Height          =   300
            Index           =   4
            Left            =   360
            TabIndex        =   23
            Top             =   2340
            Width           =   3975
         End
         Begin VB.Label lblLieferant 
            Height          =   300
            Index           =   3
            Left            =   360
            TabIndex        =   22
            Top             =   1980
            Width           =   3975
         End
         Begin VB.Label lblLieferant 
            Caption         =   "Tel-Lieferant"
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   21
            Top             =   1440
            Width           =   2535
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
            Caption         =   "IDF-Apotheke"
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame fmeAuftrag 
         Caption         =   "Auftrag"
         Height          =   4335
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5055
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
            TabIndex        =   15
            Top             =   480
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
            Index           =   1
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   14
            Top             =   1200
            Width           =   735
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
            TabIndex        =   13
            Top             =   2160
            Width           =   3655
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
            TabIndex        =   12
            Top             =   2640
            Width           =   3655
         End
         Begin VB.CheckBox chkAuftrag 
            Caption         =   "&Rückmeldungen anzeigen"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   3480
            Width           =   4335
         End
         Begin VB.CheckBox chkAuftrag 
            Caption         =   "Absagen &berücksichtigen"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   3840
            Width           =   4335
         End
         Begin VB.ComboBox cboAuftrag 
            Height          =   360
            Index           =   0
            Left            =   3240
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.ComboBox cboAuftrag 
            Height          =   360
            Index           =   1
            Left            =   3240
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   2280
            Width           =   1695
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
            TabIndex        =   17
            Top             =   480
            Width           =   3495
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
            TabIndex        =   16
            Top             =   1200
            Width           =   3495
         End
      End
   End
   Begin VB.PictureBox picSenden 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   5640
      Left            =   0
      ScaleHeight     =   5640
      ScaleWidth      =   10695
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdEscSend 
         Caption         =   "ESC"
         Height          =   450
         Left            =   6840
         TabIndex        =   44
         Top             =   3240
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Timer tmrSenden 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   2280
      End
      Begin VB.Frame fmeStatus 
         Caption         =   "Sendestatus"
         Height          =   1695
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   3495
         Begin VB.PictureBox picStatus 
            BorderStyle     =   0  'Kein
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   615
            TabIndex        =   40
            Top             =   840
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid flxAuftrag 
            Height          =   540
            Left            =   840
            TabIndex        =   41
            Top             =   960
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   953
            _Version        =   393216
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
            TabIndex        =   43
            Top             =   360
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
            Left            =   1680
            TabIndex        =   42
            Top             =   360
            Width           =   7695
         End
      End
      Begin ComctlLib.ImageList imgSenden 
         Left            =   2520
         Top             =   2400
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
               Picture         =   "wbestk2.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "wbestk2.frx":0624
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "wbestk2.frx":093E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "wbestk2.frx":0C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "wbestk2.frx":0F72
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuInfo 
         Caption         =   "&Info ..."
      End
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


''''''''''''''''''''''
'Const INI_DATEI = "\user\winop.ini"

Const INI_SECTION = "Bestellung"
Const INI_TOOLBAR_SECTION = "Wbestk2"
Const INFO_SECTION = "Infobereich Wbestk2"


Dim InRowColChange%

Dim StartZeit&

Dim HochfahrenAktiv%

Dim TimerIndex%(5)
Dim TimerStatus%

Dim AutoDirLiefs$

Private Const DefErrModul = "wbestk2.frm"
''''''''''''''''''''''

Sub Uebergabe(NeuProgrammChar$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Uebergabe")
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
Dim h$, i%, KeyStr$, ArtikelZeileKey$, ArtikelZeileVal$

UebergabeStr$ = ""

Call DefErrPop
End Sub

Public Sub WechselModus(NeuerModus%, Optional NeuMachen% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WechselModus")
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
Dim i%, j%, k%, l%, ind%, HatEingabe%, FeldInd%, f%, row%
Dim h$, h2$, h3$, mName$, ActMenuName$, OrgFontName$, OrgFontSize%

'If (Me.WindowState = vbMinimized) Or (Me.WindowState = vbMaximized) Then Me.WindowState = vbNormal
'Me.SetFocus
'DoEvents
'Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
Call FensterImVordergrund

Select Case ProgrammModus%
    Case RUFZEITENANZEIGE
        picRufzeiten.Visible = False
        tmrAction.Enabled = False
'        Me.WindowState = vbMinimized
    Case SENDEHINWEIS
        picSendGlobal.Visible = False
        picHinweis.Visible = False
        tmrHinweis.Enabled = False
    Case SENDEANZEIGE
        If (seriell% And (comSenden.PortOpen)) Then comSenden.PortOpen = False
        picSendGlobal.Visible = False
        picSenden.Visible = False
        tmrSenden.Enabled = False
    Case DIREKTBEZUG
        picSendGlobal.Visible = False
        picHinweis.Visible = False
        tmrHinweis.Enabled = False
End Select
        
ProgrammModus% = NeuerModus%

Select Case NeuerModus%
    Case RUFZEITENANZEIGE
        Me.Width = picRufzeiten.Width
        Me.Height = picRufzeiten.Height + wpara.FrmMenuHeight + 90 + wpara.FrmCaptionHeight
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2

        Call GeplanteRufzeiten
        tmrAction.Enabled = True
        picRufzeiten.Visible = True
        Me.WindowState = vbMinimized
        Call SetWindowPos(Me.hWnd, 1, 0, 0, 0, 0, 3)
        
    Case SENDEHINWEIS
        IstDirektLief% = False
        Call SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
        If (BestVorsKomplett%) Then frmBestVors.Show 1
        Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
        
        If (SchwellwertAktiv% And SchwellwertVorab%) Then
            InitSchwellwertAction
            DoEvents
            Call CheckSchwellLieferanten
            Call EinlesenLieferantenAnrufe
            Call AuslesenSchwellwertArtikel
            Close #SchwellProt%
            picSchwellwertAction.Visible = False
        End If
        
        InitSendGlobal
        InitHinweis
        
        Me.Width = picHinweis.Width
        Me.Height = picHinweis.Top + picHinweis.Height + wpara.FrmMenuHeight + 90 + wpara.FrmCaptionHeight
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2

        picSendGlobal.Visible = True
        picHinweis.Visible = True
        tmrHinweis.Enabled = True
        cmdHinweisOk.SetFocus
        
        If (Dir("wwbereit.wav") <> "") Then
            Call PlaySound("\user\wwbereit.wav", 0, SND_FILENAME Or SND_ASYNC)
        End If
    
    Case SENDEANZEIGE
        InitSendGlobal
        InitSenden
        
        Me.Width = picSenden.Width
        Me.Height = picSenden.Top + picSenden.Height + wpara.FrmMenuHeight + 90 + wpara.FrmCaptionHeight
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2

        picSendGlobal.Visible = True
        picSenden.Visible = True
        tmrSenden.Enabled = True
        flxAuftrag.SetFocus
        Call AutomaticAction
        
    Case DIREKTBEZUG
        AutomaticInd% = MAX_RUFZEITEN
        Rufzeiten(AutomaticInd%).Lieferant = Lieferant%
        Rufzeiten(AutomaticInd%).Aktiv = "J"
        Rufzeiten(AutomaticInd%).AuftragsErg = "ZH"
        Rufzeiten(AutomaticInd%).AuftragsArt = "  "

        InitSendGlobal
        InitHinweis
        
        Me.Width = picHinweis.Width
        Me.Height = picHinweis.Top + picHinweis.Height + wpara.FrmMenuHeight + 90 + wpara.FrmCaptionHeight
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2

        picSendGlobal.Visible = True
        picHinweis.Visible = True
        tmrHinweis.Enabled = True
        lstDirektBezug.SetFocus
        
        If (Dir("wwbereit.wav") <> "") Then
            Call PlaySound("\user\wwbereit.wav", 0, SND_FILENAME Or SND_ASYNC)
        End If
        
End Select

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    Select Case X \ Screen.TwipsPerPixelX
         Case WM_LBUTTONDOWN
            Me.WindowState = vbNormal
    End Select
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

If (UnloadMode = 1) Then
    Call ProgrammEnde
Else
    Cancel = 1
End If
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

Call ProgrammEnde
Call DefErrPop
End Sub

Private Sub mnuInfo_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuInfo_Click")
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
frmAbout.Show
Call DefErrPop
End Sub

Private Sub tmrAction_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrAction_Timer")
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
Dim i%, Max%, erg%, ret%, iLief%, ind%, DirektKontrollStatus%, nRufzeit%
Static TimerPos%
Dim h$, IniHinweis$

tmrAction.Enabled = False

TimerPos% = TimerPos% + 1
   
Call CheckBestzusa

Call CheckBekart

AutomaticInd% = PruefeRufzeiten%
If (AutomaticInd% >= 0) Then
    If (AutomaticInd% >= 100) Then
        AutomaticInd% = AutomaticInd% - 100
        Call WechselModus(SENDEHINWEIS)
    Else
        If (SchwellwertAktiv%) And (ManuellSendung% = False) Then
'            If (Me.WindowState = vbMinimized) Then Me.WindowState = vbNormal
'            Me.SetFocus
'            DoEvents
'            Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
            Call FensterImVordergrund
            InitSchwellwertAction
            DoEvents
            Call CheckSchwellLieferanten
            iLief% = Rufzeiten(AutomaticInd%).Lieferant
            Call EinlesenLieferantenAnrufe(AutomaticInd%)
            Call AuslesenSchwellwertArtikel
            Close #SchwellProt%
            picSchwellwertAction.Visible = False
            Me.WindowState = vbMinimized
            Call SetWindowPos(Me.hWnd, 1, 0, 0, 0, 0, 3)
        End If
        
        Lieferant% = Rufzeiten(AutomaticInd%).Lieferant
        Call AuslesenBestellung(False)
        If (AnzBestellArtikel% > 0) Then
            If (IstDirektLief% = 0) Or (lifzus.DirektBestModemKz) Then
                AutomaticSend% = True
                ret% = ModemAktivieren%
                If (ret%) Then
                    Call WechselModus(SENDEANZEIGE)
                    Call DefErrPop: Exit Sub
                Else
                    Call MsgBox("Modem momentan belegt !")
                End If
            Else
                Call InitSendGlobal
                BlindBestellung% = True
                Call SucheSendeArtikel
                Call SaetzeVorbereiten
                Call SaetzeSenden
                Call UpdateBekartDat(Lieferant%, True)
                BlindBestellung% = False
        
                lifzus.GetRecord (Lieferant% + 1)
                lifzus.TempBevorratungsZeitraum = 0
                lifzus.TempValutaStellung = 0
                lifzus.TempFakturenRabatt = 0
                lifzus.PutRecord (Lieferant% + 1)
            
                Call DirektBezugAusdruck(False)
            End If
        End If
        Call WechselModus(RUFZEITENANZEIGE)
    End If
    Call DefErrPop: Exit Sub
ElseIf (AutoDirLiefs$ <> "") Then
    If (PruefeDirektBezugSendefenster%) Then
        IstDirektLief% = True
        ind% = InStr(AutoDirLiefs$, ",")
        Lieferant% = Val(Left$(AutoDirLiefs$, ind% - 1))
        AutoDirLiefs$ = Mid$(AutoDirLiefs$, ind% + 1)
'        If (Me.WindowState = vbMinimized) Or (Me.WindowState = vbMaximized) Then Me.WindowState = vbNormal
'        Me.SetFocus
'        DoEvents
'        Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
        Call FensterImVordergrund
        InitSchwellwertAction
        DoEvents
        Call ZeigeSchwellwertAction("Einlesen Direktbezug für Lieferant: " + Format(Lieferant%, "0"))
        
        IniHinweis$ = HoleIniString$("DirektBezugHinweis")
        h$ = Format(Lieferant%, "000") + "-"
        ind% = InStr(IniHinweis$, h$)
        If (ind% <= 0) Then
            Call ZeigeSchwellwertAction("Einlesen Bestellvorschlag für Lieferant: " + Format(Lieferant%, "0"))
            DirektBezugsKz% = 1
            frmBestVors.Show 1
            DirektBezugsKz% = 0
        End If
        
        For i% = 0 To 1
            erg% = AuslesenDirektBezug(i%, DirektKontrollStatus%)
            If (erg%) Then
                picSchwellwertAction.Visible = False
                Me.WindowState = vbMinimized
                Call SetWindowPos(Me.hWnd, 1, 0, 0, 0, 0, 3)
                Call SpeicherDirektBezugAction
                Call WechselModus(DIREKTBEZUG)
                Call DefErrPop: Exit Sub
            ElseIf (DirektKontrollStatus% > 0) Then
                Exit For
            End If
        Next i%
        Call SpeicherDirektBezugAction
        picSchwellwertAction.Visible = False
        Me.WindowState = vbMinimized
        Call SetWindowPos(Me.hWnd, 1, 0, 0, 0, 0, 3)
    End If
End If

If (TimerPos% = 4) Then
    TimerPos% = 0
    If (BestVorsPeriodisch%) Then
        Call BestellVorschlag
    End If
End If

tmrAction.Enabled = True

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
Dim i%, spBreite%
Dim l&
Dim h$

On Error Resume Next

HochfahrenAktiv% = True
   
Height = wpara.WorkAreaHeight
Width = 10980 * wpara.BildFaktor '10200
    
'With picSave
'    .Left = 0
'    .Top = 0
'    .Width = ScaleWidth
'    .Height = ScaleHeight
'    .ZOrder 0
'End With


Call wpara.InitEndSub(Me)


Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI)
Call wpara.InitFont(Me)
Call HoleIniWerte

ProgrammTyp% = 0
Caption = ProgrammNamen$(ProgrammTyp%) + " - "

With flxRufzeiten
    .Cols = 10
    .Rows = 2
    .FixedRows = 1
    .FormatString = "^Lieferant|^Rufzeit|^Lieferzeit|^Wochentag||^Erg|^Art|<Aktiv||"
    .Rows = 1

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    .HighLight = flexHighlightNever
    .BackColorBkg = wpara.FarbeArbeit
    
    Call InitFlexSpalten
End With

With picRufzeiten
    .Left = 0
    .Top = 0
    .Width = flxRufzeiten.Left + flxRufzeiten.Width + 2 * wpara.LinksX%
    .Height = flxRufzeiten.Top + flxRufzeiten.Height + 90
    .Visible = False
End With

picSendGlobal.Visible = False
picHinweis.Visible = False

picRufzeiten.BackColor = wpara.FarbeArbeit
picSendGlobal.BackColor = wpara.FarbeArbeit
picHinweis.BackColor = wpara.FarbeArbeit
picSenden.BackColor = wpara.FarbeArbeit

Set SendeForm = Me

Me.Width = picRufzeiten.Width
Me.Height = picRufzeiten.Height + wpara.FrmMenuHeight + 90 + wpara.FrmCaptionHeight
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Me.WindowState = vbNormal

'Me.WindowState = vbMinimized
HochfahrenAktiv% = False

Call DefErrPop

End Sub

Sub aGeplanteRufzeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("aGeplanteRufzeiten")
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
Dim i%, k%, ind%, l%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

With flxRufzeiten
    .Redraw = False
    Call InitFlexSpalten
    .Rows = 1
    
    IstWoTag% = WeekDay(Now, vbMonday)
    IstZeit% = Val(Format(Now, "HHMM"))
    IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)
    
    IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))
    
    For k% = 1 To 7
        For i% = 0 To (AnzRufzeiten% - 1)
            If (Rufzeiten(i%).Lieferant > 0) Then
                For j% = 0 To 6
                    rWoTag% = Rufzeiten(i%).WoTag(j%)
                    rZeit% = Rufzeiten(i%).RufZeit
                    rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
                    If (rWoTag% > 0) Then
                        If (rWoTag% = IstWoTag%) And (rZeit% >= IstZeit%) Then
                            Call ZeigeGeplanteRufzeit(i%, k%, IstWoTag%)
                        End If
                    Else
                        Exit For
                    End If
                Next j%
            End If
        Next i%
    
        If (IstWoTag% = 7) Then
            IstWoTag% = 1
        Else
            IstWoTag% = IstWoTag% + 1
        End If
        IstZeit% = 6
    Next k%
    
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = " "
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = " "
        .TextMatrix(1, 3) = "Derzeit keine Rufzeiten gespeichert !"
    End If
    .row = 1
    .Col = 4
    .RowSel = .Rows - 1
    .ColSel = 4
    .Sort = 5
    .Col = 0
    .ColSel = .Cols - 1
    
    .Redraw = True
End With

Call DefErrPop
End Sub

Sub GeplanteRufzeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GeplanteRufzeiten")
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
Dim i%, k%, ind%, l%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%
Dim IstDatum&

With flxRufzeiten
    .Redraw = False
    Call InitFlexSpalten
    .Rows = 1
    
    IstWoTag% = WeekDay(Now, vbMonday)
    IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))
    
    For k% = 1 To 7
        For i% = 0 To (AnzRufzeiten% - 1)
            If (Rufzeiten(i%).Lieferant > 0) Then
                For j% = 0 To 6
                    rWoTag% = Rufzeiten(i%).WoTag(j%)
                    If (rWoTag% > 0) Then
                        If (rWoTag% = IstWoTag%) And ((k% <> 1) Or (IstDatum& <> Rufzeiten(i%).LetztSend)) Then
                            Call ZeigeGeplanteRufzeit(i%, k%, IstWoTag%)
                        End If
                    Else
                        Exit For
                    End If
                Next j%
            End If
        Next i%
    
        If (IstWoTag% = 7) Then
            IstWoTag% = 1
        Else
            IstWoTag% = IstWoTag% + 1
        End If
    Next k%
    
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = " "
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = " "
        .TextMatrix(1, 3) = "Derzeit keine Rufzeiten gespeichert !"
    End If
    .row = 1
    .Col = 4
    .RowSel = .Rows - 1
    .ColSel = 4
    .Sort = 5
    .Col = 0
    .ColSel = .Cols - 1
    
    .Redraw = True
End With

Call DefErrPop
End Sub

Sub ZeigeGeplanteRufzeit(RufInd%, SortInd%, TagInd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeGeplanteRufzeit")
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
Dim i%, k%, ind%, l%
Dim h$, h2$

lif.GetRecord (Rufzeiten(RufInd%).Lieferant + 1)

'Get #LDATEI%, Rufzeiten(RufInd%).Lieferant + 1, lif
h$ = lif.kurz
Call OemToChar(h$, h$)
h$ = h$ + " (" + Mid$(Str$(Rufzeiten(RufInd%).Lieferant), 2) + ")"

h$ = h$ + vbTab + Format$(Rufzeiten(RufInd%).RufZeit \ 100, "00")
h$ = h$ + ":" + Format$(Rufzeiten(RufInd%).RufZeit Mod 100, "00")

h$ = h$ + vbTab + Format$(Rufzeiten(RufInd%).LieferZeit \ 100, "00")
h$ = h$ + ":" + Format$(Rufzeiten(RufInd%).LieferZeit Mod 100, "00")

h$ = h$ + vbTab
'h2$ = Str$(SortInd%) + Format$(Rufzeiten(RufInd%).RufZeit, "0000")
h2$ = Format(SortInd%, "0") + Format$(Rufzeiten(RufInd%).RufZeit, "0000")
h$ = h$ + WochenTag$(TagInd% - 1)
h$ = h$ + vbTab + h2$
h$ = h$ + vbTab + Rufzeiten(RufInd%).AuftragsErg
h$ = h$ + vbTab + Rufzeiten(RufInd%).AuftragsArt
If (Rufzeiten(RufInd%).Aktiv = "J") Then
    h$ = h$ + vbTab + "ja"
Else
    h$ = h$ + vbTab + "nein"
End If
h$ = h$ + vbTab + Str$(Rufzeiten(RufInd%).Lieferant) + vbTab + Str$(RufInd%)

flxRufzeiten.AddItem h$

Call DefErrPop
End Sub

Function PruefeRufzeiten%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRufzeiten%")
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
Dim i%, j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))

Call HoleIniRufzeiten
Call GeplanteRufzeiten

ManuellSendung% = False
RueckKaufSendung% = False

For i% = 0 To (AnzRufzeiten% - 1)
    If (Rufzeiten(i%).Lieferant > 0) Then
        For j% = 0 To 6
            rWoTag% = Rufzeiten(i%).WoTag(j%)
            rZeit% = Rufzeiten(i%).RufZeit
            rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
            If (rWoTag% > 0) Then
                If (rWoTag% = IstWoTag%) And (IstDatum& <> Rufzeiten(i%).LetztSend) Then
'                    If (rZeit% = IstZeit%) And (Rufzeiten(i%).Gewarnt = "J") Then
                    If (IstZeit% >= rZeit%) And (IstZeit% <= rZeit% + AnzMinutenVerspaetung%) And (Rufzeiten(i%).Gewarnt = "J") Then
                        Rufzeiten(i%).LetztSend = IstDatum&
                        Rufzeiten(i%).Gewarnt = "N"
                        Call SpeicherIniRufzeiten
                        PruefeRufzeiten% = i%
                        Call DefErrPop: Exit Function
'                    ElseIf (IstZeit% >= rZeit% - AnzMinutenWarnung%) And (IstZeit% <= rZeit%) And (Rufzeiten(i%).Gewarnt <> "J") Then
                    ElseIf (IstZeit% >= rZeit% - AnzMinutenWarnung%) And (IstZeit% <= rZeit% + AnzMinutenVerspaetung%) And (Rufzeiten(i%).Gewarnt <> "J") Then
                        Rufzeiten(i%).Gewarnt = "J"
                        Call SpeicherIniRufzeiten
                        PruefeRufzeiten% = i% + 100
                        Call DefErrPop: Exit Function
                    End If
                End If
            Else
                Exit For
            End If
        Next j%
    End If
Next i%

For i% = 0 To (AnzRufzeiten% - 1)
    If (Rufzeiten(i%).Lieferant > 0) Then
        If (Rufzeiten(i%).RufZeit > 9000) And (Rufzeiten(i%).LetztSend = 0) Then
            Rufzeiten(i%).LetztSend = IstDatum&
            Rufzeiten(i%).Gewarnt = "N"
            Call SpeicherIniRufzeiten
            PruefeRufzeiten% = i%
            ManuellSendung% = True
            If (Rufzeiten(i%).RufZeit = 9998) Then
                RueckKaufSendung% = True
            Else
                RueckKaufSendung% = False
            End If
            Call DefErrPop: Exit Function
        End If
    End If
Next i%

PruefeRufzeiten% = -1

Call DefErrPop
End Function

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
Dim i%, cmdToolSize%, lblStatusSize%, versatz%
Dim c As Control

On Error Resume Next

If (HochfahrenAktiv%) Then Call DefErrPop: Exit Sub

If (Me.WindowState = vbMinimized) Then Call DefErrPop: Exit Sub

'Call opBereich.ResizeWindow

'picSave.Visible = False

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
Dim i%, ind%, iVal%
Dim l&, f&
Dim h$, h2$, key$
    
With frmAction
    
    iVal% = (Screen.Width - Width) / 2
    If (iVal% < 0) Then
        iVal% = 0
    End If
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "StartX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    iVal% = Val(h$)
    If (iVal% = -9999) Then
        WindowState = vbMaximized
    Else
        If (iVal% < 0) Then
            iVal% = 0
        End If
        .Left = iVal%
        
        iVal% = 75
        h$ = Format(iVal%, "00000")
        l& = GetPrivateProfileString(UserSection$, "StartY", h$, h$, 6, INI_DATEI)
        h$ = Left$(h$, l&)
        iVal% = Val(h$)
        If (iVal% < 0) Then
            iVal% = 0
        End If
        .Top = iVal%
        
        iVal% = .Width
        h$ = Format(iVal%, "00000")
        l& = GetPrivateProfileString(UserSection$, "BreiteX", h$, h$, 6, INI_DATEI)
        h$ = Left$(h$, l&)
        iVal% = Val(h$)
        If (iVal% < 0) Then
            iVal% = 0
        End If
        If (.Left + iVal% > wpara.WorkAreaWidth) Then
            WindowState = vbMaximized
        Else
            .Width = iVal%
        
            iVal% = .Height
            h$ = Format(iVal%, "00000")
            l& = GetPrivateProfileString(UserSection$, "HoeheY", h$, h$, 6, INI_DATEI)
            h$ = Left$(h$, l&)
            iVal% = Val(h$)
            If (iVal% < 0) Then
                iVal% = 0
            End If
            If (.Top + iVal% > wpara.WorkAreaHeight) Then
                WindowState = vbMaximized
            Else
                .Height = iVal%
            End If
        End If
    End If
    
''    h$ = Space$(20)
''    h2$ = Left$("Arial" + Space$(20), 20)
''    l& = GetPrivateProfileString(UserSection$, "Fontname", h2$, h$, 21, WINWAWI_INI)
''    h$ = Left$(h$, l&)
''    fName$ = RTrim$(h$)
''
''    h$ = "8"
''    l& = GetPrivateProfileString(UserSection$, "Fontsize", "8", h$, 3, WINWAWI_INI)
''    h$ = Left$(h$, l&)
''    fSize% = Val(h$)
'
'    h2$ = Left$("Arial,8" + Space$(20), 20)
'    h$ = Space$(20)
'    l& = GetPrivateProfileString(UserSection$, "FontInformation", h2$, h$, 21, WINWAWI_INI)
'    h$ = Left$(h$, l&)
'    ind% = InStr(h$, ",")
'    OpFonts(0).Name = RTrim$(Left$(h$, ind% - 1))
'    OpFonts(0).Size = Val(Mid$(h$, ind% + 1))
'
'    h2$ = Left$("Arial,8" + Space$(20), 20)
'    h$ = Space$(20)
'    l& = GetPrivateProfileString(UserSection$, "FontBeschriftung", h2$, h$, 21, WINWAWI_INI)
'    h$ = Left$(h$, l&)
'    ind% = InStr(h$, ",")
'    OpFonts(1).Name = RTrim$(Left$(h$, ind% - 1))
'    OpFonts(1).Size = Val(Mid$(h$, ind% + 1))
'
'    Call InitFont(Me)
'
'
'
'    h2$ = "C0C0C0"
'    h$ = Space$(8)
'    l& = GetPrivateProfileString(UserSection$, "FarbeArbeit", h2$, h$, 9, WINWAWI_INI)
'    h$ = Left$(h$, l&)
'    FarbeArbeit& = BerechneFarbWert&(h$)
'
'    h2$ = "80FFFF"
'    h$ = Space$(8)
'    l& = GetPrivateProfileString(UserSection$, "FarbeInfo", h2$, h$, 9, WINWAWI_INI)
'    h$ = Left$(h$, l&)
'    FarbeInfo& = BerechneFarbWert&(h$)
'
'    h$ = "0"
'    l& = GetPrivateProfileString(UserSection$, "FarbeRahmen", "0", h$, 2, WINWAWI_INI)
'    h$ = Left$(h$, l&)
'    FarbeRahmen% = Val(h$)
'
'    Call InitAlleBereichsFarben
    
    
    h$ = "01"
    l& = GetPrivateProfileString(INI_SECTION, "MinutenWarnung", "3", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    AnzMinutenWarnung% = Val(h$)
    
    h$ = "01"
    l& = GetPrivateProfileString(INI_SECTION, "MinutenWarten", "3", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    AnzMinutenWarten% = Val(h$)
    
    h$ = "01"
    l& = GetPrivateProfileString(INI_SECTION, "MinutenVerspaetung", "3", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    AnzMinutenVerspaetung% = Val(h$)
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "BestVorsKomplett", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        BestVorsKomplett% = True
    Else
        BestVorsKomplett% = False
    End If
    
    h$ = "03"
    l& = GetPrivateProfileString(INI_SECTION, "BestVorsKomplettMinuten", "03", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    BestVorsKomplettMinuten% = Val(h$)
    
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "BestVorsPeriodisch", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        BestVorsPeriodisch% = True
    Else
        BestVorsPeriodisch% = False
    End If
    
    h$ = "10"
    l& = GetPrivateProfileString(INI_SECTION, "BestVorsPeriodischMinuten", "10", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    BestVorsPeriodischMinuten% = Val(h$)

    
    h$ = "J"
    l& = GetPrivateProfileString(INI_SECTION, "Etiketten", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        MacheEtiketten% = True
    Else
        MacheEtiketten% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "AlleEtiketten", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        MacheAlleEtiketten% = True
    Else
        MacheAlleEtiketten% = False
    End If
    
    h$ = "J"
    l& = GetPrivateProfileString(INI_SECTION, "TeilDefekte", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        TeilDefekte% = True
    Else
        TeilDefekte% = False
    End If
    
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "F70Debug", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        iVal% = True
    Else
        iVal% = False
    End If
    para.F70Debug = iVal%
    
    
    h$ = "03"
    l& = GetPrivateProfileString(INI_SECTION, "IsdnEndeDelay", "03", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    IsdnEndeDelay% = Val(h$)
    
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertAktiv", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        SchwellwertAktiv% = True
    Else
        SchwellwertAktiv% = False
    End If
    
    h$ = "120"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertMinuten", "120", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    SchwellwertMinuten% = Val(h$)

    h$ = "005"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertSicherheit", "005", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    SchwellwertSicherheit% = Val(h$)

    h$ = "005"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertToleranz", "005", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    SchwellwertToleranz% = Val(h$)
    
    h$ = "100"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertWarnungProz", "100", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    SchwellwertWarnungProz% = Val(h$)

    h$ = "005"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertWarnungAb", "005", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    SchwellwertWarnungAb% = Val(h$)

    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertVorab", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        SchwellwertVorab% = True
    Else
        SchwellwertVorab% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertGlaetten", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        SchwellwertGlaetten% = True
    Else
        SchwellwertGlaetten% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString("Allgemein", "Shuttle", "N", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ShuttleAktiv% = True
    Else
        ShuttleAktiv% = False
    End If
    
    h$ = "3"
    l& = GetPrivateProfileString(INI_SECTION, "HandShake", "3", h$, 2, INI_DATEI)
    PharmaBoxHandShake% = Val(Left$(h$, l&))
    
'    h$ = "N"
'    l& = GetPrivateProfileString(INI_SECTION, "DirektBezug", "N", h$, 2, INI_DATEI)
'    h$ = Left$(h$, l&)
'    If (h$ = "J") Then
'        DirektBezugAktiv% = True
'    Else
'        DirektBezugAktiv% = False
'    End If
'    DirektBezugAktiv% = True
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "PharmaboxInDOS", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ModemInDOS% = True
    Else
        ModemInDOS% = False
    End If
    
    h$ = Space$(50)
    l& = GetPrivateProfileString(INI_SECTION, "AutomatikDrucker", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
'    If (h$ = "") Then h$ = Printer.DeviceName
    AutomatikDrucker$ = h$

    
    h$ = "10"
    l& = GetPrivateProfileString("DirektBezug", "MinimalSendefenster", "10", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    DirektBezugSendMinuten% = Val(h$)
    
    h$ = "600"
    l& = GetPrivateProfileString("DirektBezug", "MaxDauerKontrollen", "600", h$, 4, INI_DATEI)
    h$ = Left$(h$, l&)
    DirektBezugKontrollenMinunten% = Val(h$)
    
    h$ = "10"
    l& = GetPrivateProfileString("DirektBezug", "AlleMinutenWarnung", "10", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    DirektBezugWarnungMinuten% = Val(h$)
    
    
    h$ = "J"
    l& = GetPrivateProfileString(INI_SECTION, "MitLagerstand", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        MitLagerstandCalc% = True
    Else
        MitLagerstandCalc% = False
    End If

    h$ = "J"
    l& = GetPrivateProfileString(INI_SECTION, "WuPruefung", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        WuPruefung% = True
    Else
        WuPruefung% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "PartnerTeilBestellungen", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        PartnerTeilBestellungen% = True
    Else
        PartnerTeilBestellungen% = False
    End If

    h$ = Space$(50)
    l& = GetPrivateProfileString(INI_SECTION, "AbsagenMitNL", h$, h$, 51, INI_DATEI)
    AbsagenMitNL$ = Left$(h$, l&)

    h$ = Space$(50)
    l& = GetPrivateProfileString(INI_SECTION, "AutomatenLieferanten", h$, h$, 51, INI_DATEI)
    AutomatenLiefs$ = Trim(Left$(h$, l&))
    If (AutomatenLiefs$ <> "") Then
        AutomatenLac$ = Left$(h$, 1)
        AutomatenLiefs$ = Mid$(AutomatenLiefs$, 2)
    End If

End With

Call DefErrPop
End Sub

Function PruefeIniRueckrufe%(iLieferant%, Optional RRmodus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeIniRueckrufe%")
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
Dim ret%, ind%
Dim l&
Dim SollRR$, h$

ret% = False

h$ = Space$(100)
l& = GetPrivateProfileString(INI_SECTION, "Rueckrufe", " ", h$, 101, INI_DATEI)
SollRR$ = Trim(Left$(h$, l&))
    
h$ = Format(iLieferant%, "000")
ind% = InStr(SollRR$, h$)
If (ind% > 0) Then
    ret% = True
    If (RRmodus%) Then
        SollRR$ = Left$(SollRR$, ind% - 1) + Mid$(SollRR$, ind% + 3)
        Do
            ind% = InStr(SollRR$, ",,")
            If (ind% > 0) Then
                SollRR$ = Left$(SollRR$, ind% - 1) + Mid$(SollRR$, ind% + 1)
            Else
                Exit Do
            End If
        Loop
        l& = WritePrivateProfileString(INI_SECTION, "Rueckrufe", SollRR$, INI_DATEI)
    End If
End If

PruefeIniRueckrufe% = ret%

Call DefErrPop
End Function

Sub AuslesenBestellung(BereitsGelockt%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenBestellung")
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
Dim i%, Max%, AltLief%, ind%, WasGeändert%, erg%, IstLeerPzn%, nRufzeit%, iLief%
Dim zr!
Dim EK#, fr#, ZeilenWert#
Dim ManuellDirLiefs$
Dim h$, h2$, LeerPzn$, LeerTxt$, tx$, IniHinweis$, wwOrg$, wwNeu$, SQLStr$
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

'tmrAction.Enabled = False

If (BereitsGelockt% = False) Then Call ww.SatzLock(1)

ww.GetRecord (1)

IstAktiv% = True

WasGeändert% = False
Max% = ww.erstmax
AltLief% = ww.erstlief
BekartCounter% = ww.erstcounter
GlobBekMax% = Max%

AktUhrzeit% = Val(Format(Now, "HHMM"))
NaechstLiefernderLieferant% = HoleNaechstLiefernden%

zTabelleAktiv% = True
zManuellAktiv% = False

lstSortierung.Clear

If (Lieferant% > 0) And (Lieferant% <= lif.AnzRec) Then
    lif.GetRecord (Lieferant% + 1)
    h2$ = RTrim$(lif.Name(0))
    Call OemToChar(h2$, h2$)

    nRufzeit% = HoleNaechsteRufzeit%(Lieferant%)
    If (nRufzeit% >= 0) Then
        h2$ = h2$ + "  (" + Format(nRufzeit% \ 100, "00") + ":" + Format(nRufzeit% Mod 100, "00") + ")"
    End If

    Me.Caption = ProgrammNamen$(ProgrammTyp%) + " - " + h2$
End If



lifzus.GetRecord (Lieferant% + 1)

'IstDirektLief% = False
'If (DirektBezugAktiv%) Then IstDirektLief% = lifzus.IstDirektLieferant
IstDirektLief% = lifzus.IstDirektLieferant



lstSortierung.Clear
lstDirektSortierung.Clear


AnzBestellArtikel% = 0
AddMarkWert# = 0#
AddGesamtWert# = 0#

DirektBezugFaktRabatt# = 0#
DirektBezugValutaStellung% = 0
If (IstDirektLief%) Then
    DirektBezugFaktRabatt# = lifzus.FakturenRabatt
    If (lifzus.TempFakturenRabatt > 0) Then DirektBezugFaktRabatt# = lifzus.TempFakturenRabatt
    DirektBezugFaktRabattTyp% = lifzus.FakturenRabattTyp
    DirektBezugValutaStellung% = lifzus.ValutaStellung
    If (lifzus.TempValutaStellung > 0) Then DirektBezugValutaStellung% = lifzus.TempValutaStellung
    DirektBezugAbBm% = lifzus.DirektMindestBM
End If

If (ManuellSendung%) Then
    If (RueckKaufSendung%) Then

        rk.GetRecord (1)
        Max% = rk.erstmax
        
        For i% = 1 To Max%
            rk.GetRecord (i% + 1)
        
            If (rk.status = 1) And (rk.loesch = 0) And (rk.aktivlief = Lieferant%) And (rk.aktivind < 0) Then
                AnzBestellArtikel% = AnzBestellArtikel% + 1
                h$ = Left$(rk.txt, 18) + Mid$(rk.txt, 29) + Format(i%, "0000") + rk.pzn + Format(Abs(rk.bm), "0000") + Format(rk.beklaufnr, String(9, 48))
                frmAction!lstSortierung.AddItem h$
            End If
        Next i%
    Else
        For i% = 1 To Max%
            ww.GetRecord (i% + 1)
            
            If (ww.status = 1) And (ww.loesch = 0) And (ww.aktivlief = Lieferant%) And (ww.aktivind < 0) Then
                    
                AnzBestellArtikel% = AnzBestellArtikel% + 1
    
                h$ = Left$(ww.txt, 18) + Mid$(ww.txt, 29) + Format(i%, "0000") + ww.pzn + Format(Abs(ww.bm), "0000") + Format(ww.beklaufnr, String(9, 48))
                If (para.Land = "A") Then
                    h$ = h$ + ww.wg
                End If
                frmAction!lstSortierung.AddItem h$
                        
                If (IstDirektLief%) Then
                    Call InsertLstDirektSortierung
                End If
            End If
        Next i%
    End If

    If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)
    IstAktiv% = False
    Call DefErrPop: Exit Sub
End If

If (ProgrammModus% <> DIREKTBEZUG) Then
    AutoDirLiefs$ = ""
    ManuellDirLiefs$ = ""
End If

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    
    If (ww.status = 1) And (ww.loesch = 0) And (ww.aktivlief = 0) Then
        
        If (Lieferant% <> AltLief%) Then
            ww.zukontrollieren = Chr$(0)
        End If
        
        wwOrg$ = ww.RawData
        
        Call EinzelSatz(0, h$)
        
        wwNeu$ = ww.RawData
        If (wwOrg$ <> wwNeu$) Then
            ww.PutRecord (i% + 1)
            WasGeändert% = True
        End If
        
        If (ProgrammModus% <> DIREKTBEZUG) And (Abs(ww.bm) > 0) Then
            iLief% = ww.Lief
'            If (iLief% = 0) Then
'                h$ = lifzus.GetLiefFuerHerst(ww.herst)
'                If (h$ <> "") Then iLief% = Val(Left$(h$, 3))
'            End If
            
            If (iLief% > 0) Then
                lifzus.GetRecord (iLief% + 1)
                If (lifzus.IstDirektLieferant) Then
                    h$ = Format(iLief%, "000") + ","
                    If (ww.DirektTyp = 2) Then
                        If (InStr(ManuellDirLiefs$, h$) <= 0) Then
                            ManuellDirLiefs$ = ManuellDirLiefs$ + h$
                        End If
                    ElseIf (InStr(AutoDirLiefs$, h$) <= 0) And (lifzus.IstAutoDirektLieferant) Then
                        AutoDirLiefs$ = AutoDirLiefs$ + h$
                    End If
                End If
            End If
        End If
    
        If ((ww.zugeordnet = "J") And (ww.zukontrollieren <> "1")) Then
            If (ProgrammModus% <> DIREKTBEZUG) Or ((Abs(ww.bm) > 0) And (Abs(ww.bm) >= DirektBezugAbBm%)) Then
                AnzBestellArtikel% = AnzBestellArtikel% + 1
    
                h$ = Left$(ww.txt, 18) + Mid$(ww.txt, 29) + Format(i%, "0000") + ww.pzn + Format(Abs(ww.bm), "0000") + Format(ww.beklaufnr, String(9, 48))
                If (para.Land = "A") Then
                    h$ = h$ + ww.wg
                End If
                frmAction!lstSortierung.AddItem h$
                ww.aktivlief = Lieferant%
                ww.aktivind = AnzBestellArtikel%
                ww.PutRecord (i% + 1)
                WasGeändert% = True
                
                If (IstDirektLief%) Then
                    Call InsertLstDirektSortierung
                End If
            End If
        End If
    End If
Next i%

If (ProgrammModus% <> DIREKTBEZUG) Then
    Do
        ind% = InStr(ManuellDirLiefs$, ",")
        If (ind% > 0) Then
            h$ = Left$(ManuellDirLiefs$, ind%)
            ManuellDirLiefs$ = Mid$(ManuellDirLiefs$, ind% + 1)
            ind% = InStr(AutoDirLiefs$, h$)
            If (ind% > 0) Then
                h2$ = AutoDirLiefs$
                AutoDirLiefs$ = Left$(h2$, ind% - 1) + Mid$(h2$, ind% + Len(h$))
            End If
        Else
            Exit Do
        End If
    Loop
    
    
    IniHinweis$ = HoleIniString$("DirektBezugHinweis")
    If (IniHinweis$ <> "") Then
        h2$ = ""
        Do
            If (IniHinweis$ = "") Then Exit Do
            
            ind% = InStr(IniHinweis$, "-")
            If (ind% > 0) Then
                h$ = Left$(IniHinweis$, ind% - 1) + ","
                ind% = InStr(AutoDirLiefs$, h$)
                If (ind% > 0) Then
                    h2$ = h2$ + Left$(IniHinweis$, 14)
                End If
                IniHinweis$ = Mid$(IniHinweis$, 15)
            Else
                Exit Do
            End If
        Loop
        IniHinweis$ = h2$
        
        Call SpeicherIniString%("DirektBezugHinweis", IniHinweis$)
    End If
End If

LeerAuftrag% = False

If (WasGeändert%) Then
    BekartCounter% = (BekartCounter% + 1) Mod 100
    ww.erstcounter = BekartCounter%
    ww.PutRecord (1)
Else
    h2$ = lif.LeerPzn(Lieferant%)
    h2$ = RTrim$(h2$)
    If (Len(h2$) > 0) Then
        LeerAuftrag% = True
        
        Call OemToChar(h2$, h2$)
        
        IstLeerPzn% = False
        If (Len(h2$) = 7) Then
            IstLeerPzn% = True
            For i% = 1 To 7
                ind% = Asc(Mid$(h2$, i%, 1))
                If (ind% < 48) Or (ind% > 57) Then
                    IstLeerPzn% = False
                    Exit For
                End If
            Next i%
        End If
        If (IstLeerPzn%) Then
            LeerPzn$ = h2$
            LeerTxt$ = "Leerauftrag"
        Else
            LeerPzn$ = "9999999"
            LeerTxt$ = h2$
        End If
        h$ = Left$(LeerTxt$ + Space$(25), 25) + Format(0, "0000") + LeerPzn$ + Format(1, "0000") + Format(0, String(9, 48))
    '    h$ = Left$(ww.txt, 18) + Mid$(ww.txt, 29) + Format(i%, "0000") + ww.pzn + Format(Abs(ww.bm), "0000") + Format(ww.beklaufnr, String(9, 48))
        frmAction!lstSortierung.AddItem h$
        AnzBestellArtikel% = 1
    End If
End If

If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)

IstAktiv% = False

'tmrAction.Enabled = True

Call DefErrPop
End Sub

Function AuslesenDirektBezug%(AusleseModus%, DirektKontrollStatus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenDirektBezug%")
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
Dim i%, j%, Max%, AltLief%, ind%, WasGeändert%, erg%, IstLeerPzn%, AngebotY%, ret%
Dim IstBewertungOk%, NichtUeberGh%, nUeber%, ProzPlus%, iBevorratungsZeitraum%, ret2%, bm1%, mm1%
Dim DirektBewert#
Dim h$, h2$, LeerPzn$, LeerTxt$, SQLStr$, wwOrg$, wwNeu$
Static IstAktiv%

ret% = False

Call ww.SatzLock(1)

ww.GetRecord (1)

WasGeändert% = False
Max% = ww.erstmax
AltLief% = ww.erstlief
BekartCounter% = ww.erstcounter
GlobBekMax% = Max%

AktUhrzeit% = Val(Format(Now, "HHMM"))
NaechstLiefernderLieferant% = HoleNaechstLiefernden%

zTabelleAktiv% = True
zManuellAktiv% = False

If (Lieferant% > 0) And (Lieferant% <= lif.AnzRec) Then
    lif.GetRecord (Lieferant% + 1)
    lifzus.GetRecord (Lieferant% + 1)
    h2$ = RTrim$(lif.Name(0))
    Call OemToChar(h2$, h2$)
    
    DirektBezugAbBm% = lifzus.DirektMindestBM

    Call ZeigeSchwellwertAction("Prüfung Direktbezug für Lieferant: " + h2$)
End If

For j% = 0 To 1
    If (j% = 0) Then
        iBevorratungsZeitraum% = 0
    Else
        lifzus.GetRecord (Lieferant% + 1)
        ProzPlus% = lifzus.ProzentPlus
        If (ProzPlus% = 0) Then
            Exit For
        Else
            Call ZeigeSchwellwertAction("Zeitraum erhöhen um " + Format(ProzPlus%, "0") + "%")
            
            iBevorratungsZeitraum% = lifzus.TempBevorratungsZeitraum
            If (iBevorratungsZeitraum% = 0) Then iBevorratungsZeitraum% = lifzus.BevorratungsZeitraum
            If (iBevorratungsZeitraum% = 0) Then iBevorratungsZeitraum% = para.BestellPeriode
            iBevorratungsZeitraum% = Int(iBevorratungsZeitraum% * (100# + ProzPlus) / 100# + 0.501)
        
            WasGeändert% = False
            Call ww.SatzLock(1)
            ww.GetRecord (1)
            
            Max% = ww.erstmax
            BekartCounter% = ww.erstcounter
        End If
    End If

    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
        
        If (ww.status = 1) And (ww.Lief = Lieferant%) And (Abs(ww.bm) > 0) Then
        
            wwOrg$ = ww.RawData
        
            ww.zukontrollieren = Chr$(0)
            Call EinzelSatz(0, h$)
            
            wwNeu$ = ww.RawData
            If (wwOrg$ <> wwNeu$) Then
                ww.PutRecord (i% + 1)
                WasGeändert% = True
            End If
        
        End If
    Next i%
    
    DirektKontrollStatus% = 0
    Call InitDirektBewertung(Lieferant%)
    
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
        
        If (ww.status = 1) And (ww.Lief = Lieferant%) And ((Abs(ww.bm) > 0) And (Abs(ww.bm) > DirektBezugAbBm%)) Then
        
            Call EinzelSatz(1, h$, False)
            
            If ((ww.loesch = 0) And (ww.aktivlief = 0) And (ww.zugeordnet = "J")) Then
    
                Call CalcDirektBezugZeile(i%, iBevorratungsZeitraum%)
                
                Select Case ww.zukontrollieren
                    Case "1"
                        DirektKontrollStatus% = 1
                    Case "2"
                        If (DirektKontrollStatus% = 0) Then DirektKontrollStatus% = 2
                End Select
    
                WasGeändert% = True
            End If
        
        End If
    Next i%
    
    If (WasGeändert%) Then
        BekartCounter% = (BekartCounter% + 1) Mod 100
        ww.erstcounter = BekartCounter%
        ww.PutRecord (1)
    End If
    
    Call ww.SatzUnLock(1)
    
    lifzus.GetRecord (Lieferant% + 1)
    DirektBewert# = ZeigeDirektBewertung(IstBewertungOk%, AngebotY%, False)
    Call ZeigeSchwellwertAction("Bewertung Direktbezug: " + Format(DirektBewert#, "0.00"))
    
    'If (DirektBewert# < 100#) Then
    If (IstBewertungOk% = 0) Then
        Call ZeigeSchwellwertAction("Bewertung o.k.!")
        ret% = DirektBezugHinweis%(DirektKontrollStatus%)
        Exit For
    Else
        If (IstBewertungOk% And &H1) Then
            Call ZeigeSchwellwertAction("Bestellwert zu gering !")
        Else
            Call ZeigeSchwellwertAction("Bewertung zu gering !")
        End If
        
        If (j% = 0) And (IstBewertungOk% And &H1) Then
        Else
            If (AusleseModus% = 0) Then
                Call ZeigeSchwellwertAction("Löschen Bestellvorschlag !")
            Else
                Call ZeigeSchwellwertAction("eingetragenen Lieferanten bei Artikel entfernen !")
            End If
            
            WasGeändert% = False
            Call ww.SatzLock(1)
            ww.GetRecord (1)
            
            Max% = ww.erstmax
            BekartCounter% = ww.erstcounter
        
            NichtUeberGh% = False
            
            For i% = 1 To Max%
                nUeber% = 0
                
                ww.GetRecord (i% + 1)
                If (ww.status = 1) And (ww.Lief = Lieferant%) Then
                    If (AusleseModus% = 0) And (ww.DirektTyp = 1) Then
                        ww.status = 0
                        ww.PutRecord (i% + 1)
                        WasGeändert% = True
                    ElseIf (AusleseModus% = 1) Then
                        'nur wenn über GH beziehbar !
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
                        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        If (TaxeRec.EOF = False) Then
                            If (TaxeRec!UeberGH = 1) Then
                                NichtUeberGh% = True
                                nUeber% = True
                            End If
                        End If
                        If (nUeber%) Then
                        Else
                            bm1% = 0
                            FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                ass.GetRecord (FabsRecno& + 1)
                                mm1% = ass.vmm
                                If (mm1% <= 0) Then mm1% = ass.MM
                                If ((mm1% >= 1) And (ass.poslag <= Int(mm1% / 2))) Then
                                    bm1% = mm1% - ass.poslag
                                End If
                                If (mm1% = 0) And (ass.poslag = 0) Then
                                    bm1% = 1
                                End If
                            End If
                            
                            If (bm1% > 0) Then
                                ww.bm = bm1%
                                ww.nm = 0
                                ww.Lief = 0
                                ww.nnart = 0
                                ww.NNAEP = 0#
                                ww.zr = 0!
                                ww.angebot = ww.angebot And (&H34 Xor &HFFFF)
                            Else
                                ww.status = 0
                            End If
                            ww.PutRecord (i% + 1)
                            WasGeändert% = True
                        End If
                    End If
                End If
            Next i%
            
            If (WasGeändert%) Then
                BekartCounter% = (BekartCounter% + 1) Mod 100
                ww.erstcounter = BekartCounter%
                ww.PutRecord (1)
            End If
        
            Call ww.SatzUnLock(1)
            
            DirektKontrollStatus% = 0
            
            If (AusleseModus% = 1) Then
                If (NichtUeberGh%) Then
                    DirektKontrollStatus% = 3
                    ret% = DirektBezugHinweis%(DirektKontrollStatus%)
                Else
                    ret2% = DirektBezugHinweis%(DirektKontrollStatus%)
                End If
            End If
            
            Exit For
        End If
    End If
Next j%

AuslesenDirektBezug% = ret%

Call DefErrPop
End Function

Sub CalcDirektBezugZeile(pos%, iBevorratungsZeitraum%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcDirektBezugZeile")
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
Dim rInd%, angInd%, angBm%, angNm%, angLief%, angAep#, angIndOrg%, iAngebot%, angAngebotYOrg%
Dim angZr!
Dim l&
Dim h$, ch$, pzn$

Dim angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%

Dim i%, oldRow%, oldCol%, oldStatus%, AnzRows%, GesamtBewertung%, ls%, bmo!
Dim sp&
Dim ret#
Dim tx$, SQLStr$

If (iBevorratungsZeitraum% > 0) Then
    ls% = 0
    bmo! = 0
    If (ww.ssatz > 0) Then
        ls% = ww.poslag
        If (ass.opt > 0) Then bmo! = ass.opt
'        If (ass.opt > 0) Then bmo% = Int(ass.opt + 0.501)
    End If
    ww.bm = CalcDirektBM(iBevorratungsZeitraum%, bmo!, ls%)
    ww.PutRecord (pos% + 1)
End If

If (ww.bm <> 0) Then
    pzn$ = ww.pzn
        
    angBm% = Abs(ww.bm)
    angNm% = ww.nm
    angZr! = ww.zr
    
    If (ww.ssatz > 0) Then
        angBMopt! = ass.opt
        angBMast% = ass.bm
    Else
        angBMopt! = 0
        angBMast% = 0
    End If
    
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
        angTaxeAep# = TaxeRec!EK / 100
        If (ww.asatz <= 0) Then
            Call Taxe2ast(ww.pzn)
        End If
        angWgAst% = ast.wg
    Else
        angTaxeAep# = 0
        angWgAst% = 0
    End If
    
    If (ww.asatz > 0) Then
        ast.GetRecord (ww.asatz + 1)
        angAstAep# = ast.aep
    Else
        angAstAep# = 0
    End If
    
    If (ww.asatz > 0) Or (TaxeRec.EOF = False) Then
        angPznLagernd% = 1
    Else
        angPznLagernd% = 0
    End If
    
    Call RechneDirektBewertung(pzn$, angBm%, angNm%, angZr!, angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%)
End If

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
Dim l&

'Call opToolbar.SpeicherIniToolbar
'If (WindowState = vbMaximized) Then
'    l& = WritePrivateProfileString(UserSection$, "StartX", Str$(-9999), INI_DATEI)
'Else
'    l& = WritePrivateProfileString(UserSection$, "StartX", Str$(Left), INI_DATEI)
'    l& = WritePrivateProfileString(UserSection$, "StartY", Str$(Top), INI_DATEI)
'    l& = WritePrivateProfileString(UserSection$, "BreiteX", Str$(Width), INI_DATEI)
'    l& = WritePrivateProfileString(UserSection$, "HoeheY", Str$(Height), INI_DATEI)
'End If

Call KillSysTrayIcon(Me, 1)

Call DefErrPop
End Sub

Sub ErzeugeGesendeteMenu()
End Sub

Sub EndeDll()
End
End Sub

Sub FlexKurzInfo(modus%)
Static OrgKeinRowColChange%

If (modus% = 0) Then
    OrgKeinRowColChange% = KeinRowColChange%
    KeinRowColChange% = True
Else
    KeinRowColChange% = OrgKeinRowColChange%
End If

End Sub

Private Sub InitSendGlobal()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitSendGlobal")
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iLieferant%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, MitRR%
Dim h$, h2$, FormStr$, sZeit$, sAuftragsErg$


iLieferant% = Rufzeiten(AutomaticInd%).Lieferant

Call HoleLieferantenDaten(iLieferant%)

Caption = "Sendeauftrag für " + LiefName1$

MitRR% = PruefeIniRueckrufe(iLieferant%)
h$ = "ZH"
sAuftragsErg$ = Rufzeiten(AutomaticInd%).AuftragsErg
If (MitRR%) Then
    h$ = "RR"
    sAuftragsErg$ = "RR"
End If
txtAuftrag(0).text = h$

txtAuftrag(1).text = "  "
optAuftrag(0).Value = True
chkAuftrag(0).Value = 1
chkAuftrag(1).Value = 1
lblLieferant(3).Caption = LiefName1$
lblLieferant(4).Caption = LiefName2$
lblLieferant(5).Caption = LiefName3$
lblLieferant(6).Caption = LiefName4$
lblLieferantWert(0).Caption = ApoIDF$
lblLieferantWert(1).Caption = GhIDF$
lblLieferantWert(2).Caption = TelGh$
fmeAuftrag.Caption = "Auftrag"

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
        
    If (MitRR%) Then
        .ListIndex = 3
    Else
        .ListIndex = 0
    End If
    
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


'h$ = "(" + Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsErg) + Space$(2), 2) + ")"
h$ = "(" + Left$(RTrim$(sAuftragsErg$) + Space$(2), 2) + ")"
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

If (Rufzeiten(AutomaticInd%).Aktiv = "J") Then
    optAuftrag(0).Value = True
Else
    optAuftrag(1).Value = True
End If

chkAuftrag(0).Value = 0
chkAuftrag(1).Value = 1

AutomatikFertig% = False
AutomatikFehler$ = ""

fmeAuftrag.Enabled = False
For i% = 0 To 1
    lblAuftrag(i%).Enabled = False
    cboAuftrag(i%).Enabled = False
    chkAuftrag(i%).Enabled = False
    optAuftrag(i%).Enabled = False
Next i%

With picSendGlobal
    .Left = 0
    .Top = 0
    .Width = fmeLieferant.Left + fmeLieferant.Width + 2 * wpara.LinksX%
    .Height = fmeLieferant.Top + fmeLieferant.Height + 150
    .Visible = False
End With

If (ProgrammModus% = DIREKTBEZUG) Then
    fmeAuftrag.Visible = False
    
    With fmeDirektBezug
        .Left = fmeAuftrag.Left
        .Top = fmeAuftrag.Top
        .Width = fmeAuftrag.Width
        .Height = fmeAuftrag.Height
    End With
    
    With lstDirektBezug
        .Clear
        If (lifzus.DirektBestModemKz) Then .AddItem "Modem: " + Trim(lifzus.DirektBestModem)
'        If (lifzus.DirektBestMailKz) Then .AddItem "eMail: " + Trim(lifzus.DirektBestMail)
'        If (lifzus.DirektBestFaxKz) And (lifzus.DirektBestComputerFaxKz) Then .AddItem "Computer-Fax: " + Trim(lifzus.DirektBestFax)
        If (lifzus.DirektBestDruckKz) Or (lifzus.DirektBestFaxKz) Or (.ListCount = 0) Then .AddItem "Faxfähiger Ausdruck"
        
        .Height = 4 * TextHeight("Äg")
        MaxWi% = 0
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            wi% = TextWidth(.text)
            If (wi% > MaxWi%) Then MaxWi% = wi%
        Next i%
        .Width = MaxWi% + 300
        
        .Left = (fmeDirektBezug.Width - .Width) / 2
        .Top = (fmeDirektBezug.Height - .Height) / 2
        .ListIndex = 0
    
    End With

    fmeDirektBezug.Visible = True
Else
    fmeDirektBezug.Visible = False
    fmeAuftrag.Visible = True
End If

Call DefErrPop
End Sub

Private Sub InitSchwellwertAction()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitSchwellwertAction")
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

With picSchwellwertAction
    .Left = 0
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
End With
    
With flxSchwellwertAction
    .FixedRows = 0
    .Rows = 0
    .Cols = 1
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = picSchwellwertAction.Width - 2 * wpara.LinksX
    .Height = picSchwellwertAction.Height - .Top - 150
    .ColWidth(0) = .Width
End With

picSchwellwertAction.Visible = True

Call DefErrPop
End Sub

Private Sub InitHinweis()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitHinweis")
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iLieferant%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$, sZeit$

With fmeHinweis
    .Top = 0
    .Left = fmeAuftrag.Left
    .Width = fmeLieferant.Left + fmeLieferant.Width - .Left
End With


lblZeit(0).Top = 2 * wpara.TitelY%
For i% = 1 To 1
    lblZeit(i%).Top = lblZeit(i% - 1).Top + lblZeit(i% - 1).Height + 60
Next i%

lblZeit(0).Left = wpara.LinksX%
lblZeitWert(0).Top = lblZeit(0).Top
For i% = 1 To 1
    lblZeit(i%).Left = lblZeit(i% - 1).Left
    lblZeitWert(i%).Top = lblZeit(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblZeit(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblZeitWert(0).Left = lblZeit(0).Left + MaxWi% + 300
For i% = 1 To 1
    lblZeitWert(i%).Left = lblZeitWert(i% - 1).Left
Next i%


With picZeit
    .Left = wpara.LinksX%
    .Top = lblZeit(1).Top + lblZeit(1).Height + 300
    .Width = fmeHinweis.Width - .Left - wpara.LinksX%
    .Height = picHinweis.TextHeight("99 %") + 120
    fmeHinweis.Height = .Top + .Height + wpara.TitelY%
End With


cmdHinweisOk.Top = fmeHinweis.Top + fmeHinweis.Height + 150
cmdHinweisWeg.Top = cmdHinweisOk.Top


With picHinweis
    .Left = 0
    .Top = picSendGlobal.Height
    .Width = picSendGlobal.Width
    .Height = cmdHinweisOk.Top + cmdHinweisOk.Height + 90
    .Visible = False
End With


cmdHinweisOk.Width = wpara.ButtonX%
cmdHinweisOk.Height = wpara.ButtonY%
cmdHinweisWeg.Width = wpara.ButtonX%
cmdHinweisWeg.Height = wpara.ButtonY%
cmdHinweisOk.Left = (picSendGlobal.Width - (cmdHinweisOk.Width * 2 + 300)) / 2
cmdHinweisWeg.Left = cmdHinweisOk.Left + cmdHinweisWeg.Width + 300


sZeit$ = Format(Rufzeiten(AutomaticInd%).RufZeit, "0000")
lblZeitWert(0).Caption = Left$(sZeit$, 2) + ":" + Mid$(sZeit$, 3)

sZeit$ = Format(Now, "HH:MM")
lblZeitWert(1).Caption = sZeit$

h$ = Format(Now, "HHMMSS")
StartZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))

Call DefErrPop
End Sub

Private Sub tmrhinweis_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrHinweis_Timer")
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
Dim IstZeit%, rZeit%
Dim lIstZeit&, hZeit&, Warten&
Dim Prozent!
Dim sZeit$, h$

sZeit$ = Format(Now, "HH:MM")
lblZeitWert(1).Caption = sZeit$

IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)
    
h$ = Format(Now, "HHMMSS")
lIstZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))

If (ProgrammModus% = DIREKTBEZUG) Then
    Warten& = 30&
    hZeit& = lIstZeit& - StartZeit&
    Prozent! = (hZeit&) / (Warten&) * 100!
    h$ = "Verbleibende Zeit: " + Str$(Warten& - hZeit&) + " Sekunden"
Else
    Warten& = 60&
    
    rZeit% = Rufzeiten(AutomaticInd%).RufZeit
    rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
    
    hZeit& = (rZeit% * 60& - StartZeit&)
    If (hZeit& <= 0&) Then
        Prozent! = 100!
    Else
        Prozent! = (lIstZeit& - StartZeit&) / (hZeit&) * 100!
    End If
    If (lIstZeit& > rZeit% * 60&) Then
        Prozent! = 100!
    End If
    h$ = "Verbleibende Zeit: " + Str$(rZeit% * 60& - lIstZeit&) + " Sekunden"
End If

'h$ = Format$(Prozent!, "##0") + " %"
With picZeit
    .Cls
    .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
    .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
    picZeit.Print h$
    picZeit.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
'                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
End With

If (lIstZeit& - StartZeit& > Warten&) Then
    Prozent! = 100!
End If

If (rZeit% = IstZeit%) Or (Prozent! = 100!) Then
    cmdHinweisOk.Value = True
End If

Call DefErrPop
End Sub

Private Sub cmdhinweisWeg_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdHinweisWeg_Click")
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
Dim IstDatum&

If (ProgrammModus% <> DIREKTBEZUG) Then
    IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))
    Rufzeiten(AutomaticInd%).LetztSend = IstDatum&
    Rufzeiten(AutomaticInd%).Gewarnt = "N"
    Call SpeicherIniRufzeiten
End If

Call WechselModus(RUFZEITENANZEIGE)
'Unload Me
Call DefErrPop
End Sub

Private Sub cmdhinweisOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdHinweisOk_Click")
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
Dim ret%
Dim h$

If (ProgrammModus% = DIREKTBEZUG) Then
    h$ = lstDirektBezug.text
        
    ManuellSendung% = False
    RueckKaufSendung% = False
    Call AuslesenBestellung(False)
    If (AnzBestellArtikel% > 0) Then
        If (Left$(h$, 5) = "Modem") Then
            AutomaticSend% = True
            ret% = ModemAktivieren%
            If (ret%) Then
                Call WechselModus(SENDEANZEIGE)
                Call DefErrPop: Exit Sub
            Else
                Call MsgBox("Modem momentan belegt !")
            End If
'        ElseIf (Left$(h$, 5) = "eMail") Then
'            Call DirektBezugMail
'
'            BlindBestellung% = True
'            Call SucheSendeArtikel
'            Call SaetzeVorbereiten
'            Call SaetzeSenden
'            Call UpdateBekartDat(Lieferant%, True)
'            BlindBestellung% = False
'
'            Call DirektBezugAusdruck(False)
'        ElseIf (Left$(h$, 5) = "Compu") Then
        Else
            BlindBestellung% = True
            Call SucheSendeArtikel
            Call SaetzeVorbereiten
            Call SaetzeSenden
            Call UpdateBekartDat(Lieferant%, True)
            BlindBestellung% = False
    
            lifzus.GetRecord (Lieferant% + 1)
            lifzus.TempBevorratungsZeitraum = 0
            lifzus.TempValutaStellung = 0
            lifzus.TempFakturenRabatt = 0
            lifzus.PutRecord (Lieferant% + 1)
        
            Call DirektBezugAusdruck(False)
        End If
        
    End If
End If
        
Call WechselModus(RUFZEITENANZEIGE)

Call DefErrPop
End Sub

Private Sub InitSenden()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitSenden")
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

TimerIndex%(0) = 1
TimerIndex%(1) = 2
TimerIndex%(2) = 3
TimerIndex%(3) = 4
TimerIndex%(4) = 3
TimerIndex%(5) = 2

With fmeStatus
    .Top = 0
    .Left = fmeAuftrag.Left
    .Width = fmeLieferant.Left + fmeLieferant.Width - fmeStatus.Left
End With

lblModem.Left = wpara.LinksX%
lblModem.Top = 2 * wpara.TitelY%

lblModemWert.Left = lblModem.Left + lblModem.Width + 300
lblModemWert.Top = lblModem.Top
lblModemWert.Width = fmeStatus.Width - lblModemWert.Left - wpara.LinksX%


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

With cmdEscSend
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Top = fmeStatus.Top + fmeStatus.Height + 150
    .Left = (picSendGlobal.Width - .Width) / 2
End With


With picSenden
    .Left = 0
    .Top = picSendGlobal.Height
    .Width = picSendGlobal.Width
    .Height = cmdEscSend.Top + cmdEscSend.Height + 90
    .Visible = False
End With

fmeAuftrag.Caption = "Auftrag (" + Str$(AnzBestellArtikel%) + " Artikel)"
lblModemWert.Caption = ZeigeModemTyp$

fmeAuftrag.Enabled = False
fmeStatus.Visible = True
tmrSenden.Enabled = True
TimerStatus% = 0
flxAuftrag.Rows = 0
cmdEscSend.Visible = True
'cmdEscSend.Cancel = True

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

Private Sub tmrOptimal_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrOptimal_Timer")
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
tmrOptimal.Enabled = False
OptimalAus% = True
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

Sub AutomaticAction()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AutomaticAction")
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
Dim i%, erg%

If (Rufzeiten(AutomaticInd%).Aktiv = "J") Then
    AutomatikAktivSenden% = True
Else
    AutomatikAktivSenden% = False
End If
Do
        
    SendAutomatic
    If (AutomatikFertig%) Then
        If (RueckKaufSendung%) Then
            Call UpdateRueckKaufDat(Lieferant%, True)
        Else
            Call UpdateBekartDat(Lieferant%, True)
        End If
        
        If (Dir("wwfertig.wav") <> "") Then
            Call PlaySound("\user\wwfertig.wav", 0, SND_FILENAME Or SND_ASYNC)
        End If
    
        erg% = True
        Exit Do
    End If
'    Me.SetFocus
'    DoEvents
'    Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    'Call FensterImVordergrund
    frmAbbruch.Show 1
    Call FensterImVordergrund
    If (AutomatikFertig% = 2) Then
        If (RueckKaufSendung%) Then
            Call UpdateRueckKaufDat(Lieferant%, False)
        Else
            Call UpdateBekartDat(Lieferant%, False)
        End If
        erg% = False
        Exit Do
    ElseIf (AutomatikFertig% = 3) Then
        If (RueckKaufSendung%) Then
            Call UpdateRueckKaufDat(Lieferant%, True)
        Else
            Call UpdateBekartDat(Lieferant%, True)
        End If
        erg% = True
        Exit Do
    ElseIf (AutomatikFertig% = 1) Then
        AutomatikAktivSenden% = False
    Else
        AutomatikAktivSenden% = True
    End If
    AutomatikFertig% = False
Loop

If (erg% And ManuellSendung%) Then
    For i% = (AutomaticInd% + 1) To (AnzRufzeiten% - 1)
        Rufzeiten(i% - 1) = Rufzeiten(i%)
    Next i%
    AnzRufzeiten% = AnzRufzeiten% - 1
    Call SpeicherIniRufzeiten
End If

If (erg% And IstDirektLief%) Then
    lifzus.GetRecord (Lieferant% + 1)
    lifzus.TempBevorratungsZeitraum = 0
    lifzus.TempValutaStellung = 0
    lifzus.TempFakturenRabatt = 0
    lifzus.PutRecord (Lieferant% + 1)

    If (lifzus.DirektBestDruckKz) Then Call DirektBezugAusdruck(False)
End If
        
erg% = PruefeIniRueckrufe(Lieferant%, 1)
Call WechselModus(RUFZEITENANZEIGE)

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
Call DefErrPop
End Sub

Sub InitFlexSpalten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitFlexSpalten")
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
Dim i%, spBreite%

With flxRufzeiten
    Font.Bold = True
    .ColWidth(0) = TextWidth("XXXXXXXXXXXXXXX")
    .ColWidth(1) = TextWidth("99999999")
    .ColWidth(2) = TextWidth("99999999")
    .ColWidth(3) = TextWidth("XXXXXXXXXXX")
    .ColWidth(4) = 0
    .ColWidth(5) = TextWidth("XXX")
    .ColWidth(6) = TextWidth("XXX")
    .ColWidth(7) = TextWidth("XXXXXXXX")
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    Font.Bold = False

    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
End With

Call DefErrPop
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
Cancel = 0
If (CmdStr = "BEENDEN") And (ProgrammModus% = RUFZEITENANZEIGE) Then
    Call mnuBeenden_Click
End If
End Sub

Sub CheckSchwellLieferanten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckSchwellLieferanten")
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
Dim i%, erg%, rZeit%, IstZeit%, ind%, iLieferant%, iRufzeit%, iNaechstRuf%
Dim h$, IstWoTagStr$, SchwellProtName$

erg% = wpara.CreateDirectory("winw")

iLieferant% = Rufzeiten(AutomaticInd%).Lieferant
iRufzeit% = Rufzeiten(AutomaticInd%).RufZeit
iNaechstRuf% = (iRufzeit% \ 100) * 60 + (iRufzeit% Mod 100)
SchwellProtName$ = Format(iLieferant%, "000") + Format(iRufzeit, "0000")

If (Dir("winw\*.sw9") <> "") Then
    On Error Resume Next
    Kill "winw\*.sw1"
    For i% = 2 To 9
        h$ = Dir("winw\*.sw" + Format(i%, "0"))
        If (h$ <> "") Then
            h$ = "winw\" + h$
            Name h$ As Left$(h$, Len(h$) - 1) + Format(i% - 1, "0")
        End If
    Next i%
    On Error GoTo DefErr
End If

SchwellProt% = FileOpen("winw\" + SchwellProtName$ + ".sw9", "O")
   
AktSchwellLief$ = ","
AnzSchwellLief% = 0
ReDim SchwellLief(AnzSchwellLief%)

With flxRufzeiten
    IstWoTagStr$ = WochenTag$(WeekDay(Now, vbMonday) - 1)
    IstZeit% = Val(Format(Now, "HHMM"))
    IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 3) = IstWoTagStr$) Then
            rZeit% = Val(Mid$(.TextMatrix(i%, 4), 2))
            rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
'            If (SchwellwertMinuten% = 0) Or (IstZeit% + SchwellwertMinuten% >= rZeit%) Then
            If (rZeit% >= iNaechstRuf%) And ((SchwellwertMinuten% = 0) Or ((iNaechstRuf% + SchwellwertMinuten%) >= rZeit%)) Then
                h$ = Trim(.TextMatrix(i%, 0))
                ind% = InStr(h$, "(")
                If (ind% > 0) Then
                    h$ = Mid$(h$, ind% + 1)
                    h$ = Left$(h$, Len(h$) - 1)
                    iLieferant% = Val(h$)
                    
                    If (IstSchwellLieferant%(iLieferant%) < 0) Then
                        AktSchwellLief$ = AktSchwellLief$ + h$ + ","
                        ReDim Preserve SchwellLief(AnzSchwellLief%)
                        SchwellLief(AnzSchwellLief%).Lief = Val(h$)
                        SchwellLief(AnzSchwellLief%).rZeit = Val(Mid$(.TextMatrix(i%, 4), 2))
                        SchwellLief(AnzSchwellLief%).rZeit60 = rZeit%
                        
                        AnzSchwellLief% = AnzSchwellLief% + 1
                    End If
                End If
            End If
        End If
    Next i%
End With

Call ZeigeSchwellwertAction("Betroffene Lieferanten: " + AktSchwellLief$)

Call DefErrPop
End Sub

Sub ZeigeSchwellwertAction(h$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeSchwellwertAction")
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
flxSchwellwertAction.AddItem h$
DoEvents
Call DefErrPop
End Sub

Sub SpeicherDirektBezugAction()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherDirektBezugAction")
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
Dim i%, erg%, DirektBezugProt%
Dim h$, DirektbezugProtName$


erg% = wpara.CreateDirectory("winw")

If (Dir("winw\*.db9") <> "") Then
    On Error Resume Next
    Kill "winw\*.db1"
    For i% = 2 To 9
        h$ = Dir("winw\*.db" + Format(i%, "0"))
        If (h$ <> "") Then
            h$ = "winw\" + h$
            Name h$ As Left$(h$, Len(h$) - 1) + Format(i% - 1, "0")
        End If
    Next i%
    On Error GoTo DefErr
End If

DirektbezugProtName$ = Format(Lieferant%, "000") + Format(Now, "HHMM")
DirektBezugProt% = FileOpen("winw\" + DirektbezugProtName$ + ".db9", "O")
   
With flxSchwellwertAction
    For i% = 0 To (.Rows - 1)
        Print #DirektBezugProt%, .TextMatrix(i%, 0)
    Next i%
End With

Close #DirektBezugProt%

Call DefErrPop
End Sub

Function DirektBezugHinweis%(DirektKontrollStatus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugHinweis%")
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
Dim ind%, ret%, ErstZeit%, JetztZeit%
Dim h$, IniHinweis$

ret% = False

If (DirektKontrollStatus% = 1) Then
    Call ZeigeSchwellwertAction("Zwingende Kontrollen enthalten!")
ElseIf (DirektKontrollStatus% = 2) Then
    Call ZeigeSchwellwertAction("Optionale Kontrollen enthalten!")
End If

IniHinweis$ = HoleIniString$("DirektBezugHinweis")
    
h$ = Format(Lieferant%, "000") + "-"
ind% = InStr(IniHinweis$, h$)
If (ind% > 0) Then
    If (DirektKontrollStatus% = 0) Then
        IniHinweis$ = Left$(IniHinweis$, ind% - 1) + Mid$(IniHinweis$, ind% + 14)
        ret% = True
    ElseIf (DirektKontrollStatus% = 2) Then
        ErstZeit% = Val(Mid$(IniHinweis$, 5, 4))
        ErstZeit% = (ErstZeit% \ 100) * 60 + (ErstZeit% Mod 100)
        JetztZeit% = Val(Format(Now, "HHMM"))
        JetztZeit% = (JetztZeit% \ 100) * 60 + (JetztZeit% Mod 100)
    
        If (JetztZeit% < ErstZeit%) Then JetztZeit% = JetztZeit% + 24 * 60
        If ((JetztZeit% - ErstZeit%) >= DirektBezugKontrollenMinunten%) Then       '720
            IniHinweis$ = Left$(IniHinweis$, ind% - 1) + Mid$(IniHinweis$, ind% + 14)
            Call ZeigeSchwellwertAction("Hinweiszeit für Kontrollen abgelaufen!")
            ret% = True
        End If
    End If
ElseIf (DirektKontrollStatus% > 0) Then
    IniHinweis$ = IniHinweis$ + h$ + Format(Now, "HHMM") + Format(DirektKontrollStatus%, "0") + "0000,"
Else
    ret% = True
End If

Call SpeicherIniString%("DirektBezugHinweis", IniHinweis$)

If (ret%) Then Call ZeigeSchwellwertAction("Durchführen Direktbezug!")

DirektBezugHinweis% = ret%

Call DefErrPop
End Function

Sub FensterImVordergrund()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FensterImVordergrund")
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

If (Me.WindowState = vbMinimized) Or (Me.WindowState = vbMaximized) Then Me.WindowState = vbNormal
Me.SetFocus
DoEvents
Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)

Call DefErrPop
End Sub

