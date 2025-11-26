VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmWuPreisKalk 
   Caption         =   "Bestellung"
   ClientHeight    =   7890
   ClientLeft      =   -210
   ClientTop       =   675
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wupreiskalk.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   7890
   ScaleWidth      =   11295
   Begin VB.PictureBox picQuittieren 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picToolbar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      ScaleHeight     =   300
      ScaleWidth      =   1275
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      Begin VB.CommandButton cmdToolbar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picZusatzBack 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   8760
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      Begin VB.PictureBox picZusatzSymbol 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   375
         TabIndex        =   8
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
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
      Index           =   0
      Left            =   0
      ScaleHeight     =   8640
      ScaleWidth      =   10695
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdEsc 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   450
         Index           =   0
         Left            =   5280
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1200
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   450
         Index           =   0
         Left            =   3600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   3960
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   6985
         _Version        =   65541
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
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   1500
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   2646
         _Version        =   65541
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfoZusatz 
         Height          =   780
         Index           =   0
         Left            =   480
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5040
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1376
         _Version        =   65541
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblArbeit 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C0FF&
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
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   195
         Width           =   9615
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C0FF&
         Caption         =   "Information zu Artikel"
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
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   9615
      End
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":09EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":0D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":1022
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":12B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":15CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":18E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":21AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":2440
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":2964
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":2E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":33AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":363E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":38D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":3B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":3E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":4196
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":4DE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":5A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":5B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":5DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":5EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":620A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":649C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":672E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":6A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":6D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":707C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":718E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":7420
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":76B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":7944
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":7A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":7B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":7DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":808C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":819E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":8430
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":8542
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":885C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":8B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "wupreiskalk.frx":8EC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuBearbeiten 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Er&fassen"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Wechsel"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Quittieren"
         Index           =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Entfernen"
         Index           =   3
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Drucken"
         Index           =   4
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Aktualisieren"
         Index           =   5
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Zusatzinfo"
         Index           =   6
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Ab&melden"
         Index           =   7
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   9
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   10
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   11
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   12
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   13
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   14
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Rundung"
         Index           =   15
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   16
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuDummy11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBearbeitenLayout 
         Caption         =   "La&yout editieren"
      End
   End
   Begin VB.Menu mnuAnsicht 
      Caption         =   "&Ansicht"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Symbolleiste"
         Begin VB.Menu mnuToolbarVisible 
            Caption         =   "&Ausblenden"
         End
         Begin VB.Menu mnuToolbarPosition 
            Caption         =   "&Position"
            Begin VB.Menu mnuToolbarPositionInd 
               Caption         =   "&Oben"
               Checked         =   -1  'True
               Index           =   0
            End
            Begin VB.Menu mnuToolbarPositionInd 
               Caption         =   "&Rechts"
               Index           =   1
            End
            Begin VB.Menu mnuToolbarPositionInd 
               Caption         =   "&Unten"
               Index           =   2
            End
            Begin VB.Menu mnuToolbarPositionInd 
               Caption         =   "&Links"
               Index           =   3
            End
         End
         Begin VB.Menu mnuToolbarGross 
            Caption         =   "&Grosse Symbole"
         End
         Begin VB.Menu mnuToolbarLabels 
            Caption         =   "&Unterschriften"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDummy8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZusatzInfo 
         Caption         =   "Artikel-S&tatistik"
      End
   End
End
Attribute VB_Name = "frmWuPreisKalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Preiskalk"
Const INFO_SECTION = "Infobereich Preiskalk"


Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

'Dim InRowColChange%
Dim HochfahrenAktiv%
Dim ProgrammModus%
    
Dim ArtikelStatistik%

Private Const DefErrModul = "WUPREISKALK.FRM"

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
Dim i%, ind%
Dim h$

Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        
        mnuBearbeitenInd(MENU_F2).Enabled = True
        mnuBearbeitenInd(MENU_F3).Enabled = True
        mnuBearbeitenInd(MENU_F4).Enabled = True
        mnuBearbeitenInd(MENU_F5).Enabled = True
        mnuBearbeitenInd(MENU_F6).Enabled = True
        mnuBearbeitenInd(MENU_F7).Enabled = True
        mnuBearbeitenInd(MENU_F8).Enabled = True
        mnuBearbeitenInd(MENU_F9).Enabled = True
        mnuBearbeitenInd(MENU_SF2).Enabled = True
        mnuBearbeitenInd(MENU_SF3).Enabled = True
        mnuBearbeitenInd(MENU_SF4).Enabled = True
        mnuBearbeitenInd(MENU_SF5).Enabled = True
        mnuBearbeitenInd(MENU_SF6).Enabled = True
        mnuBearbeitenInd(MENU_SF7).Enabled = True
        mnuBearbeitenInd(MENU_SF8).Enabled = True
        
        mnuBearbeitenLayout.Checked = False
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        flxarbeit(0).BackColorSel = vbHighlight
        flxInfo(0).BackColorSel = vbHighlight
        
        h$ = Me.Caption
        ind% = InStr(h$, " (EDITIER-MODUS)")
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        Me.Caption = h$
    Case 1
        mnuDatei.Enabled = False
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = False
        
        mnuBearbeitenInd(MENU_F2).Enabled = True
        mnuBearbeitenInd(MENU_F3).Enabled = False
        mnuBearbeitenInd(MENU_F4).Enabled = False
        mnuBearbeitenInd(MENU_F5).Enabled = True
        mnuBearbeitenInd(MENU_F6).Enabled = False
        mnuBearbeitenInd(MENU_F7).Enabled = False
        mnuBearbeitenInd(MENU_F8).Enabled = True
        mnuBearbeitenInd(MENU_F9).Enabled = False
        mnuBearbeitenInd(MENU_SF2).Enabled = False
        mnuBearbeitenInd(MENU_SF3).Enabled = False
        mnuBearbeitenInd(MENU_SF4).Enabled = False
        mnuBearbeitenInd(MENU_SF5).Enabled = False
        mnuBearbeitenInd(MENU_SF6).Enabled = False
        mnuBearbeitenInd(MENU_SF7).Enabled = False
        mnuBearbeitenInd(MENU_SF8).Enabled = False
        
        mnuBearbeitenLayout.Checked = True
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        flxarbeit(0).BackColorSel = vbMagenta
        flxInfo(0).BackColorSel = vbMagenta
        
        h$ = Me.Caption
        Me.Caption = h$ + " (EDITIER-MODUS)"
End Select

For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%

'Me.Caption = ProgrammName$ + lblArbeit(NeuerModus%).Caption
        
ProgrammModus% = NeuerModus%

Call DefErrPop
End Sub

Private Sub cmdEsc_Click(Index As Integer)
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

Call WuPreisKalkAlle
Unload Me

Call DefErrPop
End Sub

Private Sub cmdToolbar_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdToolbar_Click")
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
If (Index = 0) Then
'    Me.WindowState = vbMinimized
ElseIf (Index <= 8) Then
    Call mnuBearbeitenInd_Click(Index - 1)
ElseIf (Index <= 16) Then
    Call mnuBearbeitenInd_Click(Index)
ElseIf (Index = 19) Then
    Call mnuBeenden_Click
End If

Call DefErrPop
End Sub

Private Sub flxarbeit_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxarbeit_DragDrop")
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
Call opToolbar.Move(flxarbeit(Index), picBack(Index), Source, X, Y)
Call DefErrPop
End Sub

Private Sub flxArbeit_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_KeyPress")
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

If (Index = 0) Then
    If (KeyAscii = vbKeySpace) Then
        Call SetzePreisZeile(flxarbeit(0).row)
'        Call NaechsteBestellZeile
    ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
        Call SelectZeile(UCase(Chr$(KeyAscii)))
    End If
End If

Call DefErrPop
End Sub

Private Sub flxArbeit_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_DblClick")
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
cmdOk(Index).Value = True
Call DefErrPop
End Sub

Private Sub flxarbeit_LostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_LostFocus")
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

Private Sub flxInfo_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxInfo_DblClick")
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
If (Index = 0) Then
    cmdOk(0).Value = True
End If

Call DefErrPop
End Sub

Private Sub flxInfo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxInfo_DragDrop")
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
Call opToolbar.Move(flxInfo(Index), picBack(Index), Source, X, Y)
Call DefErrPop
End Sub

Private Sub flxInfo_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxInfo_GotFocus")
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
Dim i%, ArbeitRow%, InfoRow%, aRow%, aCol%
Dim h$

If (Index = 0) Then
    ArbeitRow% = flxarbeit(0).row
    
    With flxInfo(0)
        aRow% = .row
        aCol% = .col
        InfoRow% = 0
        
'        Call ActProgram.flxInfoGotFocus(InfoRow%)
        
        For i% = InfoRow% To (.Rows - 1)
            .TextMatrix(i%, 0) = ""
        Next i%
        For i% = 0 To (.Rows - 1)
            .row = i%
            .col = 0
            .CellFontBold = True
            .CellForeColor = .ForeColor
        Next i%
        
        .row = aRow%
        .col = aCol%
    End With
End If
Call DefErrPop
End Sub

Private Sub flxarbeit_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_GotFocus")
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

'If (KeinRowColChange% = False) Then
    Call FormKurzInfo
'End If
Call DefErrPop
End Sub

Private Sub flxarbeit_RowColChange(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxarbeit_RowColChange")
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
Dim h$

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (Index = 0) Then
'    If ((ProgrammChar$ = "B") And (flxarbeit(0).redraw = True) And (KeinRowColChange% = False)) Then
    If ((flxarbeit(0).redraw = True) And (HochfahrenAktiv% = False)) Then
'        Call HighlightZeile
        Call FormKurzInfo
        flxInfo(0).row = 0
        flxInfo(0).col = 0
        
        picQuittieren.Visible = False
        If (flxarbeit(0).col = 2) Then
            With picQuittieren
                .Font.Name = wpara.FontName(0)
                .Font.Size = wpara.FontSize(0)
                .Width = flxarbeit(0).Width
                .Height = flxarbeit(0).Height
                .Cls
                .CurrentY = 90
                .CurrentX = 90
                h$ = flxarbeit(0).TextMatrix(flxarbeit(0).row, 2)
                picQuittieren.Print h$
                .Width = TextWidth(h$) + 300
                .Height = .CurrentY + 150
                .Top = picBack(0).Top + flxarbeit(0).Top + flxarbeit(0).RowPos(flxarbeit(0).row) + flxarbeit(0).RowHeight(0)
                .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(2)
                .Visible = True
            End With
        End If
    End If
End If
    
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
Dim i%
Dim l&
Dim h$

HochfahrenAktiv% = True

Width = Screen.Width - (800 * wpara.BildFaktor)
Height = Screen.Height - (1200 * wpara.BildFaktor)
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Caption = "Preiskalkulation"

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
    .Visible = True
End With

Set opToolbar = New clsToolbar
Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION)

cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
cmdToolbar(1).ToolTipText = "F2 Alphatext-Eingabe"
cmdToolbar(2).ToolTipText = "F3 Umschalten der Anzeige"
cmdToolbar(3).ToolTipText = "F4"
cmdToolbar(4).ToolTipText = "F5 Entfernen"
cmdToolbar(5).ToolTipText = "F6"
cmdToolbar(6).ToolTipText = "F7"
cmdToolbar(7).ToolTipText = "F8 Zusatztext"
cmdToolbar(8).ToolTipText = "F9 Abmelden"
cmdToolbar(9).ToolTipText = "shift+F2 Bestell-Status"
cmdToolbar(10).ToolTipText = "shift+F3"
cmdToolbar(11).ToolTipText = "shift+F4"
cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
cmdToolbar(13).ToolTipText = "shift+F6"
cmdToolbar(14).ToolTipText = "shift+F7 Rundung"
cmdToolbar(15).ToolTipText = "shift+F8"
cmdToolbar(16).ToolTipText = "shift+F9"
'cmdToolbar(19).ToolTipText = "Programm beenden"

cmdToolbar(5).Enabled = False
mnuBearbeitenInd(MENU_F6).Enabled = cmdToolbar(5).Enabled



Call wpara.InitFont(Me)
Call HoleIniWerte

Set InfoMain = New clsInfoBereich
Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, INFO_SECTION)
Call InfoMain.ZeigeInfoBereich("", False)
flxInfo(0).row = 0
flxInfo(0).col = 0

Set opBereich = New clsOpBereiche
Call opBereich.InitBereich(Me, opToolbar)
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = False
opBereich.ArbeitWasDarunter = False
opBereich.InfoTitel = False
opBereich.InfoZusatz = ArtikelStatistik%
opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = -2

mnuZusatzInfo.Checked = ArtikelStatistik%

ProgrammModus% = 0

With flxarbeit(0)
    .Cols = 24
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>AEP|>AVP|>rund.AVP|>kalk.AVP|<Kalkulation|>tAVP|>POS|^v|>Wg|A"
    .Rows = 1
    Call ActProgram.PreisKalkBefuellen
    For i% = 1 To (.Rows - 1)
        Call PruefeAvpRot(i%)
    Next i%
    .SelectionMode = flexSelectionFree
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 5
    .RowSel = .Rows - 1
    .ColSel = 6
    .CellFontBold = True
    
    .row = 1
    .col = 7
    .RowSel = .Rows - 1
    .ColSel = 9
    .CellBackColor = vbWhite
    
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 23) = "*") Then
            .row = i%
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellForeColor = vbBlue
        End If
    Next i%
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 5
End With
        
HochfahrenAktiv% = False
picBack(0).Visible = True

Call DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lblarbeit_DragDrop")
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
Call opToolbar.Move(lblArbeit(Index), picBack(Index), Source, X, Y)
Call DefErrPop
End Sub

Private Sub lblInfo_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lblInfo_DragDrop")
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
Call opToolbar.Move(lblInfo(Index), picBack(Index), Source, X, Y)
Call DefErrPop
End Sub

Private Sub mnuBearbeitenInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBearbeitenInd_Click")
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
Dim i%, erg%, row%, col%
Dim l&
Dim h$, mErg$

Select Case Index

    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call FormKurzInfo
            End If
        End If
    
    Case MENU_F5
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call FormKurzInfo
            End If
        End If
        
    Case MENU_F6
        Call ActProgram.KontrollListe(1)
    
    Case MENU_F7
        With flxarbeit(0)
            .redraw = False
            row% = .row
            For i% = 1 To .Rows - 1
                Call SetzePreisZeile(i%, 1)
            Next i%
            .row = row%
            .redraw = True
        End With
        
    Case MENU_F8
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                col% = flxInfo(0).col
                If (col% > 0) And (col% Mod 2) Then
                    row% = flxInfo(0).row
                    If (InfoMain.Bezeichnung(row%, (col% - 1) \ 2) <> "") Then
                        Call EditInfoName
                    End If
                End If
            End If
        End If
    
    Case MENU_SF2
'        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF3
'        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF4
'        Call ActProgram.MenuBearbeiten(Index)
            
    Case MENU_SF5
'        Call ZeigeStatistik
        
    Case MENU_SF6
'        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF7
        Call ToggleRundung

    Case MENU_SF8
'        Call ActProgram.MenuBearbeiten(Index)

    Case MENU_SF9
End Select

Call DefErrPop
End Sub

Private Sub mnuBearbeitenLayout_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBearbeitenLayout_Click")
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

If (ProgrammModus% = 0) Then
    Call WechselModus(1)
Else
    Call WechselModus(0)
End If

Call DefErrPop
End Sub

Private Sub mnuZusatzinfo_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuZusatzInfo_Click")
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
Dim h$

If (mnuZusatzInfo.Checked) Then
    mnuZusatzInfo.Checked = False
Else
    mnuZusatzInfo.Checked = True
End If

opBereich.InfoZusatz = mnuZusatzInfo.Checked

If (opBereich.InfoZusatz) Then
    h$ = "J"
Else
    h$ = "N"
End If
l& = WritePrivateProfileString(INI_SECTION, "ArtikelStatistik", h$, INI_DATEI)

opBereich.RefreshBereich

Call DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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
Dim i%, ind%, h$

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (Shift And vbCtrlMask And (KeyCode <> 17)) Then
    ind% = -1
    Select Case KeyCode
        Case vbKeyF2
            ind% = 1
        Case vbKeyF3
            ind% = 2
        Case vbKeyF4
            ind% = 3
        Case vbKeyF5
            ind% = 4
        Case vbKeyF6
            ind% = 5
        Case vbKeyF7
            ind% = 6
        Case vbKeyF8
            ind% = 7
        Case vbKeyF9
            ind% = 8
'        Case vbKeyF11
'            ind% = 9
    End Select
    If ((Shift And vbShiftMask) And (ind% > 0)) Then
        ind% = ind% + 8
    End If
    If (ind% >= 0) Then
        h$ = cmdToolbar(ind%).ToolTipText
        picToolTip.Width = picToolTip.TextWidth(h$ + "x")
        picToolTip.Height = picToolTip.TextHeight(h$) + 45
        picToolTip.Left = picToolbar.Left + cmdToolbar(ind%).Left
        picToolTip.Top = picToolbar.Top + picToolbar.Height + 60
        picToolTip.Visible = True
        picToolTip.Cls
        picToolTip.CurrentX = 2 * Screen.TwipsPerPixelX
        picToolTip.CurrentY = 0
        picToolTip.Print h$
        KeyCode = 0
    End If
End If

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
Dim i%, cmdToolSize%, lblStatusSize%, versatz%
Dim c As Control

On Error Resume Next

If (HochfahrenAktiv%) Then Call DefErrPop: Exit Sub

If (Me.WindowState = vbMinimized) Then Call DefErrPop: Exit Sub

Call opBereich.ResizeWindow
picSave.Visible = False

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
cmdEsc(0).Value = True
Call DefErrPop
End Sub

Private Sub mnuToolbarGross_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuToolbarGross_Click")
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

If (opToolbar.BigSymbols) Then
    opToolbar.BigSymbols = False
Else
    opToolbar.BigSymbols = True
End If

Call DefErrPop
End Sub

Private Sub mnuToolbarLabels_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuToolbarLabels_Click")
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
If (opToolbar.Labels) Then
    opToolbar.Labels = False
Else
    opToolbar.Labels = True
End If

Call DefErrPop
End Sub

Private Sub mnuToolbarPositionInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuToolbarPositionInd_Click")
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

opToolbar.Position = Index

Call DefErrPop
End Sub

Private Sub mnuToolbarVisible_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuToolbarVisible_Click")
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
If (opToolbar.Visible) Then
    opToolbar.Visible = False
    mnuToolbarVisible.Caption = "Einblenden"
Else
    opToolbar.Visible = True
    mnuToolbarVisible.Caption = "Ausblenden"
End If

Call DefErrPop
End Sub

Private Sub picBack_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picBack_DragDrop")
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
Call opToolbar.Move(picBack(Index), picBack(Index), Source, X, Y)
Call DefErrPop
End Sub

Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picBack_MouseMove")
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
If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

Call DefErrPop
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picToolbar_MouseDown")
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
picToolbar.Drag (vbBeginDrag)
opToolbar.DragX = X
opToolbar.DragY = Y
Call DefErrPop
End Sub

Public Sub cmdOk_Click(Index As Integer)
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
Dim row%, col%
Dim h$

If (ProgrammModus% = 1) Then
    If (ActiveControl.Name = flxInfo(0).Name) Then
        With flxInfo(0)
            row% = .row
            col% = .col
            h$ = RTrim(.text)
        End With
        If (col% Mod 2) Then
            Call InfoMain.EditInfoBelegung
            Call FormKurzInfo
        End If
    End If
ElseIf (ActiveControl.Name = flxarbeit(0).Name) Then
    Call EditSatz
ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
    With flxInfo(0)
        row% = .row
        col% = .col
        h$ = RTrim(.text)
    End With
    If (col% = 0) Then
'        Call ActProgram.cmdOkClick(h$)
    End If
End If

Call DefErrPop
End Sub

Public Sub FormKurzInfo()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FormKurzInfo")
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
Dim row%, iRow%, iCol%
Dim pzn$
    
If (flxarbeit(0).Rows < 2) Then Call DefErrPop: Exit Sub
row% = flxarbeit(0).row
pzn$ = RTrim$(flxarbeit(0).TextMatrix(row%, 0))
If (pzn$ = "XXXXXXX") Then Call DefErrPop: Exit Sub
If (Len(pzn$) <> 7) Then Call DefErrPop: Exit Sub
pzn$ = KorrPzn$(pzn$)

row% = flxarbeit(0).row
iRow% = flxInfo(0).row
iCol% = flxInfo(0).col
Call InfoMain.ZeigeInfoBereich(pzn$, True)
If (ActiveControl.Name <> flxInfo(0).Name) Then
'    Call ZeigeInfoBereichAdd(flxarbeit(0).TextMatrix(row%, 16), 0)
End If
flxInfo(0).row = iRow%
flxInfo(0).col = iCol%

If (opBereich.InfoZusatz) Then
    Call ZeigeInfoZusatz(pzn$)
End If

FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    mnuBearbeitenInd(MENU_SF5).Enabled = True
Else
    mnuBearbeitenInd(MENU_SF5).Enabled = False
End If
cmdToolbar(12).Enabled = mnuBearbeitenInd(MENU_SF5).Enabled


If (RTrim$(flxarbeit(0).TextMatrix(flxarbeit(0).row, 13)) <> "") Then
    mnuBearbeitenInd(MENU_SF4).Enabled = True
Else
    mnuBearbeitenInd(MENU_SF4).Enabled = False
End If
cmdToolbar(11).Enabled = mnuBearbeitenInd(MENU_SF4).Enabled

With flxarbeit(0)
    If (Trim(.TextMatrix(.row, 22)) = "") Then
        opToolbar.RundungFlag (False)
    Else
        opToolbar.RundungFlag (True)
    End If
End With

Call DefErrPop
End Sub

Sub ZeigeInfoZusatz(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInfoZusatz")
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
Dim i%, j%, Monat%, erg%
Dim Jahr&, AltJahr&, Termin&
Dim iWert!

With flxInfoZusatz(0)
    .redraw = False
    .GridLines = flexGridInset
    .SelectionMode = flexSelectionFree
    .Rows = 2
    .FixedRows = 1
    .Cols = 15
  
    .FillStyle = flexFillRepeat
    .col = 0
    .row = 1
    .ColSel = .Cols - 1
    .RowSel = .Rows - 1
    .CellBackColor = vbWhite
    
    .col = 0
    .row = 0
    .ColSel = .Cols - 1
    .RowSel = .Rows - 1
    .CellAlignment = flexAlignCenterCenter
    .FillStyle = flexFillSingle

    erg% = artstat.StatistikRechnen(pzn$)
    If (erg%) Then
        AltJahr& = -1
        j% = 0
        For i% = 0 To 12
            Termin& = artstat.Anfang - i% - 1
            Jahr& = (Termin& - 1) \ 12
            Monat% = Termin& - Jahr& * 12
            If (AltJahr& <> Jahr&) Then
                .TextMatrix(0, j%) = Str$(Jahr&)
                iWert! = artstat.JahresWert(Jahr&)
                If (iWert! <> 0!) Then
                    .TextMatrix(1, j%) = Str$(iWert!)
                Else
                    .TextMatrix(1, j%) = ""
                End If
                .row = 0
                .col = j%
                .CellFontBold = True
                .CellFontSize = .Font.Size
                .row = 1
                .CellBackColor = .BackColorFixed
                .CellFontBold = True
                .CellFontSize = .Font.Size
                j% = j% + 1
                AltJahr& = Jahr&
            End If
            
            .TextMatrix(0, j%) = para.MonatKurz(Monat%)
            
            iWert! = artstat.MonatsWert(i% + 1)
            If (iWert! = 0) Then
                .TextMatrix(1, j%) = ""
            Else
                .TextMatrix(1, j%) = Str$(iWert!)
            End If
            j% = j% + 1
        Next i%
    Else
        For i% = 0 To .Cols - 1
            .TextMatrix(0, i%) = ""
            .TextMatrix(1, i%) = ""
        Next i%
    End If
    
    .redraw = True
End With

Call DefErrPop

End Sub

Sub SelectZeile(SearchLetter$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SelectZeile")
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
Dim i%, gef%
Dim ch$
        
gef% = -1
With flxarbeit(0)
    For i% = (.row + 1) To (.Rows - 1)
        ch$ = Left$(.TextMatrix(i%, 2), 1)
        If (ch$ = SearchLetter$) Then
            gef% = i%
            Exit For
        End If
    Next i%
    
    If (gef% < 0) Then
        For i% = 1 To (.row - 1)
            ch$ = Left$(.TextMatrix(i%, 2), 1)
            If (ch$ = SearchLetter$) Then
                gef% = i%
                Exit For
            End If
        Next i%
    End If
        
    If (gef% > 0) Then
'        Call HighlightZeile(True)
        .row = gef%
        .col = 8
        
        If (.row < .TopRow) Then
            .TopRow = .row
        Else
            If (.row >= (.TopRow + opBereich.ArbeitAnzZeilen - 2)) Then
                .TopRow = .row - opBereich.ArbeitAnzZeilen + 2
            End If
    '        While ((.row - .TopRow) >= (ParentBereich.ArbeitAnzZeilen - 1))
    '            .TopRow = .TopRow + 1
    '        Wend
        End If
'        Call HighlightZeile
        Call FormKurzInfo
    End If
End With

Call DefErrPop
End Sub

Private Sub opToolbar_Resized()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("opToolbar_Resized")
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
Call opBereich.ResizeWindow
Call DefErrPop
End Sub

Sub RefreshBereichsFlexSpalten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshBereichsFlexSpalten")
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
Dim i%, j%, spBreite%
Dim sp&
            
With flxarbeit(0)
    Font.Bold = True
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth("X")
    .ColWidth(2) = 0
    .ColWidth(3) = TextWidth("XXXXXX")
    .ColWidth(4) = TextWidth("XXX")
    .ColWidth(5) = TextWidth("99999.99 ")
    .ColWidth(6) = TextWidth("99999.99 ")
    .ColWidth(7) = TextWidth("99999.99 ")
    .ColWidth(8) = TextWidth("99999.99 ")
    .ColWidth(9) = TextWidth("Stamm-AEP + AMPV") + wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    For i% = 10 To (.Cols - 1)
        .ColWidth(i%) = 0    'TextWidth("99999.99")
    Next i%
    Font.Bold = False
    
    
    spBreite% = 0
    For i% = 1 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(2) = .Width - spBreite%
End With
    

With flxInfo(0)
    sp& = .Width / 8
    .ColWidth(0) = 2 * sp&
    For i% = 1 To 6
        .ColWidth(i%) = sp&
    Next i%
End With
        
With flxInfoZusatz(0)
    .Cols = 15
    sp& = .Width / 15
    For i% = 0 To 14
        .ColWidth(i%) = sp&
    Next i%
End With
        
Call DefErrPop
End Sub

Sub RefreshBereichsControlsAdd()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshBereichsControlsAdd")
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

On Error Resume Next

Call DefErrPop
End Sub

Sub RefreshBereichsFarbenAdd()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshBereichsFarbenAdd")
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

On Error Resume Next

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
    
h$ = "N"
l& = GetPrivateProfileString(INI_SECTION, "ArtikelStatistik", "N", h$, 2, INI_DATEI)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    ArtikelStatistik% = True
Else
    ArtikelStatistik% = False
End If
    
Call DefErrPop
End Sub

Sub EditInfoName()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditInfoName")
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
Dim EditRow%, EditCol%
Dim h2$

EditModus% = 1

EditRow% = flxInfo(0).row
EditCol% = flxInfo(0).col

Load frmEdit

With frmEdit
    .Left = picBack(0).Left + flxInfo(0).Left + flxInfo(0).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = picBack(0).Top + flxInfo(0).Top + EditRow% * flxInfo(0).RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
    .Width = flxInfo(0).ColWidth(EditCol%)
    .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit.txtEdit
    .Width = frmEdit.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = InfoMain.Bezeichnung(EditRow%, (EditCol% - 1) \ 2)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit.Show 1
           
If (EditErg%) Then
    If (Trim$(EditTxt$) <> "") Then
        InfoMain.Bezeichnung(EditRow%, (EditCol% - 1) \ 2) = EditTxt$
        Call FormKurzInfo
    End If
End If

Call DefErrPop
End Sub

Sub ToolbarUpdateFlag(tbStatus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToolbarUpdateFlag")
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
Call opToolbar.UpdateFlag(tbStatus%)
Call DefErrPop
End Sub








Sub SetzePreisZeile(row%, Optional typ% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetzePreisZeile")
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
Dim col%

With flxarbeit(0)
    .row = row%
    col% = .col
    .col = 1
    .CellFontName = "Symbol"
    If (typ% Or (.TextMatrix(row%, 1) <> Chr$(214))) Then
        If (Trim(.TextMatrix(row%, 7)) <> "") Then
            .TextMatrix(row%, 6) = .TextMatrix(row%, 7)
            Call PruefeAvpRot(row%)
        End If
        If (KalkOhnePreis%) Or (CDbl(.TextMatrix(row%, 6)) > 0#) Then
            .TextMatrix(row%, 1) = Chr$(214)
            cmdToolbar(5).Enabled = True
            mnuBearbeitenInd(MENU_F6).Enabled = cmdToolbar(5).Enabled
        End If
    Else
        .TextMatrix(row%, 1) = " "
        .TextMatrix(row%, 6) = .TextMatrix(row%, 21)
        Call PruefeAvpRot(row%)
    End If
    .col = col%
End With

Call DefErrPop
End Sub

Sub ToggleRundung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToggleRundung")
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
Dim row%, col%
Dim h$

With flxarbeit(0)
    .redraw = False
    row% = .row
    col% = .col
    .col = 7
    If (.CellFontItalic) Then
        .CellFontItalic = False
        h$ = PruefeRundung()
        .TextMatrix(row%, 7) = Trim$(h$)
        .TextMatrix(row%, 22) = ""
        opToolbar.RundungFlag (False)
    Else
        .CellFontItalic = True
        .TextMatrix(row%, 7) = .TextMatrix(row%, 8)
        .TextMatrix(row%, 22) = "*"
        opToolbar.RundungFlag (True)
    End If
    If (Trim(.TextMatrix(row%, 7)) <> "") Then
        .TextMatrix(row%, 6) = .TextMatrix(row%, 7)
        Call PruefeAvpRot(row%)
    End If
    
    .col = col%
    .redraw = True
End With

Call DefErrPop
End Sub

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iKalk%, mw%, aufschl%, aCol%, iKalkModus%
Dim KalkAvp#
Dim KalkText$, Col1$
Dim h$, h2$
            
EditModus% = 4
            
EditCol% = flxarbeit(0).col
If (EditCol% >= 5) And (EditCol% <= 9) Then
            
    With flxarbeit(0)
        aRow% = .row
        .row = 0
        .CellFontBold = True
        .row = aRow%
    End With
            
    Load frmEdit
    
    If (EditCol% >= 7) Then
        lInd% = Val(flxarbeit(0).TextMatrix(flxarbeit(0).row, 17))
        aRow% = 0
        With frmEdit
            .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(8) + 45
            .Left = .Left + Left + wpara.FrmBorderHeight
            .Top = picBack(0).Top + flxarbeit(0).Top + flxarbeit(0).RowHeight(0)
            .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
            .Width = flxarbeit(0).ColWidth(8) + flxarbeit(0).ColWidth(9)
            .Height = flxarbeit(0).Height - flxarbeit(0).RowHeight(0)
        End With
        With frmEdit.flxEdit
            .Height = frmEdit.ScaleHeight
            frmEdit.Height = .Height
            .Width = frmEdit.ScaleWidth
            .Left = 0
            .Top = 0
            
            .Rows = 0
            .Cols = 3
            
            .ColWidth(0) = flxarbeit(0).ColWidth(8)
            .ColWidth(1) = TextWidth("Stamm-AEP + AMPV  ")
            .ColWidth(2) = .Width - .ColWidth(0) - .ColWidth(1)
            .ColAlignment(2) = flexAlignRightCenter
            
            .AddItem vbTab + "(freie Kalkulation)" + vbTab + Str$(KALK_FREIE)
            .AddItem vbTab + "(Preisempfehlung)" + vbTab + Str$(KALK_PREISEMPFEHLUNG)
            .AddItem vbTab + "(aktueller Preis)" + vbTab + Str$(KALK_AKTUELLER_PREIS)
            
            If (flxarbeit(0).TextMatrix(flxarbeit(0).row, 12) = "v") And (AufschlagsTabelle(MAX_AUFSCHLAEGE - 1).PreisBasis <> 0) Then
                aRow% = 3
                Call AvpKalkulation(MAX_AUFSCHLAEGE, KalkAvp#, KalkText$)
                h$ = ""
                If (KalkAvp# > 0#) Then h$ = Format(KalkAvp#, "0.00")
                h$ = h$ + vbTab + KalkText$ + vbTab + Str$(MAX_AUFSCHLAEGE)
                .AddItem h$
            Else
                For i% = 0 To (MAX_AUFSCHLAEGE - 2)
                    If ((i% + 1) = lInd%) Then aRow% = i% + 3
                    Call AvpKalkulation(i% + 1, KalkAvp#, KalkText$)
                    h$ = ""
                    If (KalkAvp# > 0#) Then h$ = Format(KalkAvp#, "0.00")
                    h$ = h$ + vbTab + KalkText$ + vbTab + Str$(i% + 1)
                    .AddItem h$
                Next i%
            End If
            
            
            .row = aRow%
            .col = 0
            .ColSel = .Cols - 1
            .Visible = True
        End With
    Else
        With frmEdit
            .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(EditCol%) + 45
            .Left = .Left + Left + wpara.FrmBorderHeight
            .Top = picBack(0).Top + flxarbeit(0).Top + (flxarbeit(0).row - flxarbeit(0).TopRow + 1) * flxarbeit(0).RowHeight(0)
            .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
            .Width = flxarbeit(0).ColWidth(EditCol%)
            .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
        End With
        With frmEdit.txtEdit
            .Width = frmEdit.ScaleWidth
    '            .Height = frmEdit.ScaleHeight
            .Left = 0
            .Top = 0
            h2$ = flxarbeit(0).TextMatrix(flxarbeit(0).row, EditCol%)
            .text = h2$
            .BackColor = vbWhite
            .Visible = True
        End With
    End If
   
    frmEdit.Show 1
            
    With flxarbeit(0)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
            
        If (EditErg%) Then
            If (EditCol% >= 7) Then
                m% = 0
                ind% = 0
                Do
                    ind% = InStr(ind% + 1, EditTxt$, vbTab)
                    If (ind% <= 0) Then Exit Do
                    m% = m% + 1
                Loop
                If (m% < 2) Then
                    EditTxt$ = vbTab + EditTxt$
                End If
            
                Col1$ = " "
                
                ind% = InStr(EditTxt$, vbTab)
                h$ = Left$(EditTxt$, ind% - 1)
                .TextMatrix(.row, 8) = Trim$(h$)
                EditTxt$ = Mid$(EditTxt$, ind% + 1)
                ind% = InStr(EditTxt$, vbTab)
                h$ = Left$(EditTxt$, ind% - 1)
                .TextMatrix(.row, 9) = Trim$(h$)
                EditTxt$ = Mid$(EditTxt$, ind% + 1)
                .TextMatrix(.row, 17) = Trim$(EditTxt$)
                
                iKalkModus% = Val(.TextMatrix(.row, 17))
                If (iKalkModus% = KALK_FREIE) Then
                    ManuellTxt$ = .TextMatrix(.row, 2) + " " + .TextMatrix(.row, 3) + .TextMatrix(.row, 4)
                    FreiKalkPreise#(0) = CDbl(.TextMatrix(.row, 18))
                    FreiKalkPreise#(1) = CDbl(.TextMatrix(.row, 5))
                    FreiKalkPreise#(2) = CDbl(.TextMatrix(.row, 19))
                    FreiKalkMw$ = .TextMatrix(.row, 20)
                    frmFreieKalk.Show 1
                    If (ManuellErg%) Then
                        .TextMatrix(.row, 8) = ManuellTxt$
                    Else
                        .TextMatrix(.row, 8) = ""
                    End If
                ElseIf (iKalkModus% = KALK_PREISEMPFEHLUNG) Or (iKalkModus% = KALK_AKTUELLER_PREIS) Then
                    If (iKalkModus% = KALK_PREISEMPFEHLUNG) Then
                        .TextMatrix(.row, 6) = .TextMatrix(.row, 10)
                        Call PruefeAvpRot(.row)
                        rInd% = SucheFlexZeile(True)
                        If (rInd% > 0) Then
                            Call ResetKzAufschlag
                            .TextMatrix(.row, 17) = Format(0, "0")
                        End If
                    Else
                        .TextMatrix(.row, 17) = Format(KALK_AVP_EINGABE, "0")
                    End If
                    
                    .TextMatrix(.row, 8) = ""
                    .TextMatrix(.row, 9) = "(willk. AVP-Eingabe)"
                    .TextMatrix(.row, 22) = ""
                    Col1$ = Chr$(214)
                End If
                
                If (.TextMatrix(.row, 8) = "") Then
                    .TextMatrix(.row, 7) = " "
                    .TextMatrix(.row, 8) = " "
                Else
                    If (Trim(.TextMatrix(.row, 22)) = "") Then
                        h$ = PruefeRundung()
                    Else
                        h$ = .TextMatrix(.row, 8)
                    End If
                    .TextMatrix(.row, 7) = Trim$(h$)
                    If (Trim$(h$) <> "") Then
                        .TextMatrix(.row, 6) = .TextMatrix(.row, 7)
                        Call PruefeAvpRot(.row)
                    End If
                    Col1$ = Chr$(214)
                End If
                
                aCol% = .col
                .col = 1
                .CellFontName = "Symbol"
                .col = aCol%
                .TextMatrix(.row, 1) = Col1$
            Else
                .TextMatrix(.row, EditCol%) = Format(Val(EditTxt$), "0.00")
                If (EditCol% = 6) Then
                    Call PruefeAvpRot(.row)
                    .TextMatrix(.row, 17) = Format(KALK_AVP_EINGABE, "0")
                    
                    aCol% = .col
                    .col = 1
                    .CellFontName = "Symbol"
                    .col = aCol%
                    
                    If (CDbl(.TextMatrix(.row, 6)) > 0#) Then
                        .TextMatrix(.row, 1) = Chr$(214)
                    Else
                        .TextMatrix(.row, 1) = " "
                    End If
                    .TextMatrix(.row, 7) = " "
                    .TextMatrix(.row, 8) = " "
                    .TextMatrix(.row, 9) = "(willk. AVP-Eingabe)"
                    .TextMatrix(.row, 22) = ""
                End If
            End If
            
            If (.TextMatrix(.row, 1) = Chr$(214)) Then
                cmdToolbar(5).Enabled = True
                mnuBearbeitenInd(MENU_F6).Enabled = cmdToolbar(5).Enabled
            End If
        End If
    End With
End If

Call DefErrPop
End Sub



Sub AvpKalkulation(PkInd%, KalkAvp#, KalkText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AvpKalkulation")
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
Dim iNnAep#, iStammAep#, iTaxeAep#, iPkInd%
Dim iMw$

KalkAvp# = 0#
KalkText$ = ""

iPkInd% = PkInd%
If (iPkInd% > 1000) Then iPkInd% = iPkInd% - 1000
If (iPkInd% >= 500) Then
    iPkInd% = iPkInd% - 500
    KalkText$ = "EURO-Preis"
End If

If (iPkInd% > 0) And (iPkInd% < 100) Then
    With flxarbeit(0)
        iNnAep# = CDbl(.TextMatrix(.row, 18))
        iStammAep# = CDbl(.TextMatrix(.row, 5))
        iTaxeAep# = CDbl(.TextMatrix(.row, 19))
        iMw$ = .TextMatrix(.row, 20)
        KalkAvp# = ActProgram.AvpKalkulation(iPkInd%, iNnAep#, iStammAep#, iTaxeAep#, iMw$, KalkText$)
    End With
End If

Call DefErrPop
End Sub

Function PruefeRundung$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRundung$")
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
Dim OrgAvp#, KalkAvp#, RundAvp#

With flxarbeit(0)
    OrgAvp# = CDbl(.TextMatrix(.row, 21))   '6 war falsch weil änderbar
    KalkAvp# = CDbl(.TextMatrix(.row, 8))
End With

RundAvp# = ActProgram.PruefeRundung(OrgAvp#, KalkAvp#)

If (RundAvp# > 0) Then
    PruefeRundung$ = Format(RundAvp#, "0.00")
Else
    PruefeRundung$ = ""
End If

Call DefErrPop
End Function



Sub WuPreisKalkAlle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuPreisKalkAlle")
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
Dim i%, rInd%, PkInd%
Dim NeuAep#
Dim pzn$

PreisKalkErg% = True
PreisKalkAepChange% = False

Call ww.SatzLock(1)
With flxarbeit(0)
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = Chr$(214)) Then
            .row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                NeuAep# = CDbl(.TextMatrix(i%, 5))
                If (NeuAep# <> ww.WuAEP) Then
                    ww.WuAEP = NeuAep#
                    PreisKalkAepChange% = True
                End If
                ww.WuAVP = CDbl(.TextMatrix(i%, 6))
                PkInd% = Val(.TextMatrix(i%, 17))
                If (Trim(.TextMatrix(i%, 22)) <> "") Then PkInd% = PkInd% + 1000
                If (Trim(.TextMatrix(i%, 23)) <> "") Then PkInd% = PkInd% + 500
                Call ActProgram.SpeicherKalkPreise(rInd%, PkInd%)
            End If
        Else
            PreisKalkErg% = False
        End If
    Next i%
End With
Call ww.SatzUnLock(1)

Call DefErrPop
End Sub

Function SucheFlexZeile%(Optional BereitsGelockt% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheFlexZeile%")
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
Dim i%, row%, pos%, ret%, Max%
Dim LaufNr&
Dim pzn$, ch$

ret% = False

With flxarbeit(0)
    row% = .row
    LaufNr& = Val(.TextMatrix(row%, 15))
    pos% = Val(Right$(.TextMatrix(row%, 16), 5))
    
    If (BereitsGelockt% = False) Then Call ww.SatzLock(1)
    ww.GetRecord (1)
    Max% = ww.erstmax
    
    ret% = SucheDateiZeile%(pos%, Max%, LaufNr&)
    
    If (ret%) Then
'        If (bek.aktivlief > 0) Then
'            Call MsgBox("Bestellsatz gesperrt!")
'            ret% = False
'        End If
    Else
        Call iMsgBox("WÜ-Satz nicht mehr vorhanden!")
    End If
    
    If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)
End With

SucheFlexZeile% = ret%

Call DefErrPop
End Function

Sub PruefeAvpRot(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeAvpRot")
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
Dim aRow%, aCol%
Dim iForeColor&
Dim tAvp#, iAvp#

With flxarbeit(0)
    iAvp# = CDbl(.TextMatrix(row%, 6))
    tAvp# = CDbl(.TextMatrix(row%, 10))
    If (tAvp# > 0#) Then
        aRow% = .row
        aCol% = .col
        .redraw = False
        
        .row = row%
        .col = 6
        
        If (iAvp# > tAvp#) Then
            iForeColor& = vbRed
        Else
            iForeColor& = .ForeColor
        End If
        If (.CellForeColor <> iForeColor&) Then
            .CellForeColor = iForeColor&
        End If
        
        .row = aRow%
        .col = aCol%
        .redraw = True
    End If
End With

Call DefErrPop
End Sub

Sub ResetKzAufschlag()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResetKzAufschlag")
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
Dim pzn$
                
pzn$ = ww.pzn
If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        ast.ka = "0"
        ast.PutRecord (FabsRecno& + 1)
    End If
    
    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ass.GetRecord (FabsRecno& + 1)
        ass.pk = 0
        ass.PutRecord (FabsRecno& + 1)
    End If
End If

Call DefErrPop
End Sub


