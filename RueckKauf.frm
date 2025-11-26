VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRueckKauf 
   Caption         =   "Bestellung"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RueckKauf.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   6690
   ScaleWidth      =   9660
   Begin VB.Timer tmrAction 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5520
      Top             =   240
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   10
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         _Version        =   393216
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5040
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   1376
         _Version        =   393216
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
         NumListImages   =   25
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":09EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":0B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":0C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":0EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":11BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":14D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":17F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":1B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":1D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":20B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":21CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":245C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":2776
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":2888
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":299A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":2CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":2F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":31D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":346A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":3784
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":3A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":3DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":40D2
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
         NumListImages   =   25
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":43EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":44FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":4790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":48A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":49B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":4C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":4ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":51F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":550C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":5826
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":5AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":5DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":5EE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":6176
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":6490
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":65A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":66B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":69CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":6AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":6BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":6D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":701E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":7338
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":7652
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "RueckKauf.frx":796C
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
         Caption         =   ""
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
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
         Caption         =   "Artike&l-Status"
         Index           =   9
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Lieferant"
         Index           =   10
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   11
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "S&tatistik"
         Index           =   12
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Sendevorgang"
         Index           =   13
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   14
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   15
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Rückkauf-Anfrage"
         Index           =   16
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuDummy11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBearbeitenLayout 
         Caption         =   "La&yout editieren"
      End
      Begin VB.Menu mnuDummy12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   ""
         Index           =   0
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
Attribute VB_Name = "frmRueckKauf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "RueckKauf"
Const INFO_SECTION = "Infobereich RueckKauf"


Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

'Dim InRowColChange%
Dim HochfahrenAktiv%
Dim ProgrammModus%
    
Dim ArtikelStatistik%

Dim AnzRueckKaufArtikel%
Dim RueckKaufLieferant%

Dim FrmActionTmrActionEnabled%
Dim FrmActionTmrRowaEnabled%
Dim GlobRkMax%, RkCounter%
Dim EinzelneRk%

'Dim OrgProgChar$

Private Const DefErrModul = "RUECKKAUF.FRM"

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
Static EnabledArray%(20)

Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        
        On Error Resume Next
        For i% = MENU_F2 To MENU_SF9
            mnuBearbeitenInd(i%).Enabled = EnabledArray%(i%)
        Next i%
        On Error GoTo DefErr
        
'        mnuBearbeitenInd(MENU_F2).Enabled = True
'        mnuBearbeitenInd(MENU_F3).Enabled = False
'        mnuBearbeitenInd(MENU_F4).Enabled = False
'        mnuBearbeitenInd(MENU_F5).Enabled = True
'        mnuBearbeitenInd(MENU_F6).Enabled = True
'        mnuBearbeitenInd(MENU_F7).Enabled = True
'        mnuBearbeitenInd(MENU_F8).Enabled = True
'        mnuBearbeitenInd(MENU_F9).Enabled = True
'        mnuBearbeitenInd(MENU_SF2).Enabled = True
'        mnuBearbeitenInd(MENU_SF3).Enabled = True
'        mnuBearbeitenInd(MENU_SF4).Enabled = False
'        mnuBearbeitenInd(MENU_SF5).Enabled = True
'        mnuBearbeitenInd(MENU_SF6).Enabled = True
'        mnuBearbeitenInd(MENU_SF7).Enabled = False
'        mnuBearbeitenInd(MENU_SF8).Enabled = False
        
        mnuBearbeitenLayout.Checked = False
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        flxarbeit(0).BackColorSel = vbHighlight
        flxInfo(0).BackColorSel = vbHighlight
        
        tmrAction.Enabled = True
        
        h$ = Me.Caption
        ind% = InStr(h$, " (EDITIER-MODUS)")
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        Me.Caption = h$
    Case 1
        mnuDatei.Enabled = False
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = False
        
        For i% = MENU_F2 To MENU_SF9
            EnabledArray%(i%) = mnuBearbeitenInd(i%).Enabled
        Next i%
        
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
        
        tmrAction.Enabled = False
        
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

rk.CloseDatei

With frmAction.tmrAction
     .Enabled = FrmActionTmrActionEnabled%
End With
With frmAction.tmrRowa
     .Enabled = FrmActionTmrRowaEnabled%
End With
'ProgrammChar$ = OrgProgChar$

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

If (KeinRowColChange% = False) Then
    Call FormKurzInfo
End If

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
    If ((flxarbeit(0).redraw = True) And (KeinRowColChange% = False)) Then
        Call HighlightZeile
        flxInfo(0).row = 0
        flxInfo(0).col = 0
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

With frmAction.tmrAction
    FrmActionTmrActionEnabled% = .Enabled
    .Enabled = False
End With
With frmAction.tmrRowa
    FrmActionTmrRowaEnabled% = .Enabled
    .Enabled = False
End With

'OrgProgChar$ = ProgrammChar$
'ProgrammChar$ = "W"



Width = Screen.Width - (600 * wpara.BildFaktor)
Height = Screen.Height - (1200 * wpara.BildFaktor)
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

Caption = "Rückkauf-Anfragen"

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
    .Visible = True
End With

If (ProgrammChar$ = "B") Then
    h$ = ""
Else
    h$ = "2314"
End If
Set opToolbar = New clsToolbar
Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION, h$)

cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
cmdToolbar(1).ToolTipText = "F2 Alphatext-Eingabe"
cmdToolbar(2).ToolTipText = "F3"
cmdToolbar(3).ToolTipText = "F4"
cmdToolbar(4).ToolTipText = "F5 Entfernen"
cmdToolbar(5).ToolTipText = "F6 Ausdruck"
cmdToolbar(6).ToolTipText = "F7"
cmdToolbar(7).ToolTipText = "F8 Zusatztext"
cmdToolbar(8).ToolTipText = "F9 Abmelden"
cmdToolbar(9).ToolTipText = "shift+F2 Bestell-Status"
cmdToolbar(10).ToolTipText = "shift+F3 Lieferanten-Wahl"
cmdToolbar(11).ToolTipText = "shift+F4"
cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
cmdToolbar(13).ToolTipText = "shift+F6 Sendevorgang"
cmdToolbar(14).ToolTipText = "shift+F7"
cmdToolbar(15).ToolTipText = "shift+F8"
cmdToolbar(16).ToolTipText = "shift+F9"
'cmdToolbar(19).ToolTipText = "Programm beenden"

'cmdToolbar(5).Enabled = False
'mnuBearbeitenInd(MENU_F6).Enabled = cmdToolbar(5).Enabled



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

rk.OpenDatei
If (rk.DateiLen = 0) Then
    rk.erstmax = 0
    rk.erstlief = 0
    rk.erstcounter = 0
    rk.erstrest = String(rk.DateiLen, 0)
    rk.PutRecord (1)
End If

With flxarbeit(0)
    .Cols = 12
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>RetM|^Verfall|^Lieferant|||"
    .Rows = 1
    
    RueckKaufLieferant% = 0
    RkLifDat$ = Chr$(0)
    Call EntferneGeloeschteRueckKauf
    Call AuslesenRueckKauf
    
    .SelectionMode = flexSelectionFree
    
    .row = 1
    .col = 5
End With
        

mnuBearbeitenZusatz(0).Caption = "&Blind-Bestellung"
mnuBearbeitenZusatz(0).Enabled = (ProgrammChar$ = "B")

mnuBearbeitenInd(MENU_F2).Enabled = (ProgrammChar$ = "B")
cmdToolbar(9).Enabled = mnuBearbeitenInd(MENU_F2).Enabled
    
If (ProgrammChar$ = "B") Then
    h$ = "Sendevorgang"
Else
    h$ = "Retoure erstellen"
End If
mnuBearbeitenInd(MENU_SF6).Caption = h$
cmdToolbar(13).ToolTipText = "shift+F6 " + h$

Call WechselModus(1)
Call WechselModus(0)

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
Dim i%, erg%, row%, col%, ind%, ret%, OrgLief%
Dim l&
Dim h$, mErg$, pzn$, txt$, Jetzt$, iLifDat$

tmrAction.Enabled = False
Select Case Index

    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call FormKurzInfo
            End If
        Else
            mErg$ = MatchCode(4, pzn$, txt$, False, False)
            If (mErg$ <> "") Then
                Jetzt$ = Format(Now, "DDMMYY") + Format(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)), "0000")
                Do
                    If (mErg$ = "") Then Exit Do

                    ind% = InStr(mErg$, vbTab)
                    h$ = Left$(mErg$, ind% - 1)
                    mErg$ = Mid$(mErg$, ind% + 1)

                    ind% = InStr(h$, "@")
                    pzn$ = Left$(h$, ind% - 1)
                    h$ = Mid$(h$, ind% + 1)

                    ind% = InStr(h$, "@")
                    txt$ = Left$(h$, ind% - 1)
                    h$ = Mid$(h$, ind% + 1)

                    ManuellPzn$ = pzn$
                    ManuellTxt$ = txt$
                    Call ManuellBefuellen
                    Call NeuerRueckKauf(h$)
                Loop
                Call AuslesenRueckKauf
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
        Else
            Call LoescheRueckKaufZeile
            Call NaechsteRueckKaufZeile
        End If
        
    Case MENU_F6
        Call DruckeRueckKauf
    
    Case MENU_F7
        Call EntferneGeloeschteRueckKauf
        Call AuslesenRueckKauf
        
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
        
    Case MENU_SF3
        If (ProgrammChar$ = "B") Then
            If (RueckKaufLieferant% = 0) Then
                txt$ = ""
            Else
                Call lif.GetRecord(RueckKaufLieferant% + 1)
                txt$ = UCase(Trim$(lif.kurz))
            End If
            mErg$ = MatchCode(1, pzn$, txt$, False, False)
            If (mErg$ <> "") Then
                RueckKaufLieferant% = Val(pzn$)
                Call AuslesenRueckKauf
            End If
        Else
            h$ = RkLifDat$
            WuAuswahlModus% = 1
            frmWuAuswahl.Show 1
            If (RkLifDat$ <> h$) Then
                Call AuslesenRueckKauf
            End If
        End If
        
    Case MENU_SF4
            
    Case MENU_SF5
        With flxarbeit(0)
            Call ZeigeStatbild(.TextMatrix(.row, 0), Me)
        End With
        AppActivate Me.Caption
        
    Case MENU_SF6
        If (ProgrammChar$ = "B") Then
            Call AuslesenRueckKauf2
            If (AnzBestellArtikel% > 0) Then
                If (Wbestk2ManuellSenden%) Then
                    Wbestk2ManuellVorbereitung% = True
                    AutomaticSend% = False
                    ManuellSendung% = True
                    LeerAuftrag% = False
                    RueckKaufSendung% = True
                    OrgLief% = Lieferant%
                    Lieferant% = RueckKaufLieferant%
                    frmSenden.Show 1
                    Wbestk2ManuellVorbereitung% = False
                    Lieferant% = OrgLief%
                    RueckKaufSendung% = False
    '                Call HoleIniRufzeiten
    '                i% = AnzRufzeiten%
    '                Rufzeiten(i%).Lieferant = Lieferant%
    '                Rufzeiten(i%).Aktiv = "J"
    '                Rufzeiten(i%).AuftragsErg = "ZH"
    '                Rufzeiten(i%).AuftragsArt = "  "
    '                Rufzeiten(i%).Gewarnt = "J"
    '                Rufzeiten(i%).RufZeit = 9999
    '                Rufzeiten(i%).LieferZeit = 9999
    '                Rufzeiten(i%).LetztSend = 0
    '                AnzRufzeiten% = AnzRufzeiten% + 1
    '                Call SpeicherIniRufzeiten
    '                Call HoleFruehesteManuelleSendezeit%
                    Call AuslesenRueckKauf
                Else
                    OrgLief% = Lieferant%
                    Lieferant% = RueckKaufLieferant%
                    RueckKaufSendung% = True
                    ret% = ModemAktivieren%
                    If (ret%) Then
                        AutomaticSend% = False
                        ManuellSendung% = True
                        LeerAuftrag% = False
                        frmSenden.Show 1
                    Else
                        Call iMsgBox("Modem momentan belegt !")
                    End If
                    Lieferant% = OrgLief%
                    RueckKaufSendung% = False
                End If
            End If
        Else
            Call RkFertigAlle
        End If
        
    Case MENU_SF7

    Case MENU_SF8
'        Call ActProgram.MenuBearbeiten(Index)

    Case MENU_SF9
End Select
tmrAction.Enabled = True

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
    
If (Me.Visible = False) Then Call DefErrPop: Exit Sub

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
If (ActiveControl.Name <> flxInfo(0).Name) And (ProgrammChar$ <> "B") Then
    Call ZeigeInfoBereichAdd(flxarbeit(0).TextMatrix(row%, 10), 0)
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


'If (RTrim$(flxarbeit(0).TextMatrix(flxarbeit(0).row, 13)) <> "") Then
'    mnuBearbeitenInd(MENU_SF4).Enabled = True
'Else
'    mnuBearbeitenInd(MENU_SF4).Enabled = False
'End If
'cmdToolbar(11).Enabled = mnuBearbeitenInd(MENU_SF4).Enabled

Call DefErrPop
End Sub

Sub ZeigeInfoBereichAdd(sLifDat$, Index%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInfoBereichAdd")
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
Dim j%, Lief%
Dim h$, h2$, h3$

h3$ = ""

With flxInfo(Index%)
    .redraw = False
    
    .row = 0
    .col = 0
    .CellFontBold = True
    .CellForeColor = .ForeColor
    
    Lief% = Asc(sLifDat$)
    lif.GetRecord (Lief% + 1)
    h$ = RTrim$(lif.kurz)
    h$ = h$ + " ("
    h2$ = Mid$(sLifDat$, 2, 6)
    h$ = h$ + Left$(h2$, 2) + "." + Mid$(h2$, 3, 2) + "." + Right$(h2$, 2)
    h2$ = Mid$(sLifDat$, 8, 4)  ' Format(CVI(Mid$(sLifDat$, 8, 2)), "0000")
    h$ = h$ + " " + Left$(h2$, 2) + ":" + Mid$(h2$, 3)
    h$ = h$ + ")"
    .TextMatrix(0, 0) = h$
    
'    .row = 1
'    .col = 0
'    .CellFontBold = True
'    .CellForeColor = .ForeColor
'    .CellAlignment = flexAlignLeftCenter
'    .TextMatrix(1, 0) = flxarbeit(Index%).TextMatrix(flxarbeit(Index%).row, 25)
'
'    .row = 2
'    .col = 0
'    .CellFontBold = True
'    .CellForeColor = vbBlue
'    .TextMatrix(2, 0) = h3$
'
'    j% = 4
'    Do While (j% <= ParentInfo.AnzInfoZeilen%)
'        .TextMatrix(j% - 1, 0) = ""
'        j% = j% + 1
'    Loop
    
    .redraw = True
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
    .ColWidth(5) = TextWidth("99999")
    .ColWidth(6) = TextWidth("99:9999")
    If (ProgrammChar$ = "B") Then
        .ColWidth(7) = 0
    Else
        .ColWidth(7) = TextWidth("(XXXXXX)")
    End If
    .ColWidth(8) = 0
    .ColWidth(9) = 0
    .ColWidth(10) = 0
    .ColWidth(11) = wpara.FrmScrollHeight '+ 2 * wpara.FrmBorderHeight
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
            
tmrAction.Enabled = False

EditCol% = flxarbeit(0).col
If (EditCol% >= 5) And (EditCol% <= 6) Then
            
    With flxarbeit(0)
        aRow% = .row
        .row = 0
        .CellFontBold = True
        .row = aRow%
    End With
            
    Load frmEdit
    
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
   
    If (EditCol% = 5) Then
        EditModus% = 0
    Else
        EditModus% = 5
    End If
    
    frmEdit.Show 1
            
    With flxarbeit(0)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
            
        If (EditErg%) Then
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                If (EditCol% = 5) Then
                    m% = Val(EditTxt$)
                    rk.bm = m%
                    rk.PutRecord (rInd% + 1)
                    .TextMatrix(.row, EditCol%) = Format(m%, "0")
                Else
                    h2$ = rk.WuAblDatum
                    h$ = "01" + EditTxt$
                    If (h$ <> h2$) Then
                        rk.WuAblDatum = h$
                        rk.PutRecord (rInd% + 1)
                        .TextMatrix(.row, EditCol%) = EditTxt$
                    End If
                End If
                Call ErhoeheRueckKaufCounter
            End If
        End If
    End With
End If

tmrAction.Enabled = True

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
    LaufNr& = Val(.TextMatrix(row%, 9))
    pos% = Val(Right$(.TextMatrix(row%, 8), 5))
    
    rk.GetRecord (1)
    Max% = rk.erstmax
    
    ret% = SucheRueckKaufZeile%(pos%, Max%, LaufNr&)
    
    If (ret%) Then
'        If (bek.aktivlief > 0) Then
'            Call MsgBox("Bestellsatz gesperrt!")
'            ret% = False
'        End If
    Else
        Call iMsgBox("Rückkauf-Satz nicht mehr vorhanden!")
    End If
    
End With

SucheFlexZeile% = ret%

Call DefErrPop
End Function

Sub AuslesenRueckKauf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenRueckKauf")
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
Dim i%, j%, Max%, l%, AltLief%, ind%, row%, WasGeõndert%, AnzeigeCol%, erg%, iLief%
Dim nRufzeit%, DoppeltKontrolle%, gef%, Bm0Kontrolle%, Erst%, IstOk%
Dim h$, h2$, wwOrg$, wwNeu$
Dim AutoDirLiefs$, ManuellDirLiefs$, IniHinweis$, h3$, EinzelHinweis$, LifDat$
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

tmrAction.Enabled = False

rk.GetRecord (1)

IstAktiv% = True

Max% = rk.erstmax
AltLief% = rk.erstlief
RkCounter% = rk.erstcounter
GlobRkMax% = Max%

With flxarbeit(0)
    .redraw = False
    If (.Rows < 2) Then
        .Rows = 2
    End If
    .Rows = 1
    .row = 0
End With


If (ProgrammChar$ = "B") Then
    If (RueckKaufLieferant% = 0) Then RueckKaufLieferant% = AltLief%

    h$ = "Rückkaufanfragen - Sendung an "

    If (RueckKaufLieferant% > 0) And (RueckKaufLieferant% <= lif.AnzRec) Then
        lif.GetRecord (RueckKaufLieferant% + 1)
        h2$ = RTrim$(lif.kurz)
    Else
        h2$ = "??????"
    End If
    h$ = h$ + h2$ + " "

    lblArbeit(0) = h2$

    Caption = h$
Else
    h$ = RkLifDat$
    If (Len(h$) < 3) Then
        If (Len(h$) < 2) Then EinzelneRk% = False
        h3$ = Left$(h$, 1)
        h$ = Mid$(h$, 2)
        If (h3$ = Chr$(0)) Then
            h2$ = "Alle Lieferanten"
        Else
            iLief% = Asc(h3$)
            lif.GetRecord (iLief% + 1)
            h2$ = RTrim$(lif.Name(0))
        End If
        If (h$ = "") Or (h$ = " ") Then
            h2$ = h2$ + " (alle Rückkauf-Anfragen)"
        End If
    Else
        iLief% = Asc(Mid$(h$, 2, 1))
        lif.GetRecord (iLief% + 1)
        h2$ = RTrim$(lif.Name(0))
        h2$ = h2$ + " ("
            
        Erst% = True
        Do
            If (Left$(h$, 1) = "@") Then
                LifDat$ = Mid$(h$, 2, 11)
                h$ = Mid$(h$, 14)
            
                If (Erst% = False) Then
                    h2$ = h2$ + ", "
                End If
                h3$ = Mid$(LifDat$, 2, 6)
                h2$ = h2$ + Left$(h3$, 2) + "." + Mid$(h3$, 3, 2) + "." + Right$(h3$, 2)
                h3$ = Mid$(LifDat$, 8, 4)   'Format(CVI(Right$(LifDat$, 2)), "0000")
                h2$ = h2$ + " " + Left$(h3$, 2) + ":" + Mid$(h3$, 3)
                Erst% = False
            Else
                Exit Do
            End If
        Loop
        h2$ = h2$ + ")"
        
    End If
    
    lblArbeit(0) = h2$
    
    h$ = "Gesendete Rückkaufanfragen - "
    Caption = "Gesendete Rückkaufanfragen - " + h2$

End If



AnzRueckKaufArtikel% = 0

LifDat$ = ""
For i% = 1 To Max%
    rk.GetRecord (i% + 1)

    IstOk% = False
    If (ProgrammChar$ = "B") Then
        If (rk.status = 1) Then IstOk% = True
    ElseIf (rk.status = 2) Then
        LifDat$ = Chr$(rk.Lief) + rk.WuBestDatum + Format(CVI(rk.WuBestZeit), "0000")

        IstOk% = True
        If (Len(RkLifDat$) <= 2) Then
            h3$ = Left$(RkLifDat$, 1)
            h$ = Mid$(RkLifDat$, 2)
            If (h3$ <> Chr$(0)) And (h3$ <> Left$(LifDat$, 1)) Then
                IstOk% = False
            End If
        ElseIf (Len(RkLifDat$) > 2) And (InStr(RkLifDat$, "@" + LifDat$) = 0) Then
            IstOk% = False
        ElseIf (Asc(rk.WuBestDatum) = 0) Then
            IstOk% = False
        End If

    End If



    If (IstOk%) Then
        With flxarbeit(0)
            AnzRueckKaufArtikel% = AnzRueckKaufArtikel% + 1

'            If (.row >= (.Rows - 1)) Then
                .Rows = .Rows + 1
                .row = .Rows - 1
'            End If
            Call ZeigeRueckKaufZeile(i%)
            .TextMatrix(.row, 10) = LifDat$
        End With
    End If
Next i%

If (ProgrammChar$ = "B") Then
    AltLief% = RueckKaufLieferant%
    rk.erstlief = AltLief%
    rk.PutRecord (1)
End If


'''''''''''
With flxarbeit(0)
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = "XXXXXXX"
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = "Keine anzuzeigenden Rueckkauf-Anfragen gespeichert !"
        For i% = 3 To (.Cols - 1)
            .TextMatrix(1, i%) = ""
        Next i%
    Else
        .Rows = .row + 1
    End If
    .row = 1
    .col = 8
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = 8
    .Sort = 5
    .col = 5
    .redraw = True
    
    Call HighlightZeile
End With

tmrAction.Enabled = True

Call opToolbar.UpdateFlag(False)

IstAktiv% = False

Call DefErrPop
End Sub

Sub ZeigeRueckKaufZeile(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeRueckKaufZeile")
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
Dim i%, j%, Lief%, l%, row%, col%, iBold%, iItalic%, FirstRed%, iAngebot%, red14%, bmo%, xAktiv%
Dim lBack&
Dim EK#, Rabatt#
Dim h$, h2$, s$, LiefName$, ArtName$, ArtMenge$, ArtMeh$, VonWo$, nm$, zusatz$, KontrollChar$
Dim ActKontrollen$, actzuordnung$, SRT$, SQLStr$, ZusInfo$
Dim ArtikelKz As Byte


'KeinRowColChange% = True

h$ = rk.txt
SRT$ = h$ + Format(pos%, "00000")

LiefName$ = ""
Lief% = rk.Lief
If (Lief% > 0) And (Lief% <= lif.AnzRec) Then
    lif.GetRecord (Lief% + 1)
    h2$ = lif.kurz
'    If (Trim(h2$) = "") Then h2$ = "???"
    If (Trim(h2$) = "") Or (Asc(Left$(h2$, 1)) < 32) Then h2$ = Format(Lief%, "0")
    LiefName$ = h2$
End If

If (rk.pzn = "9999999") Then
    h$ = RTrim$(rk.txt)
    Call OemToChar(h$, h$)
    ArtName$ = h$
    ArtMenge$ = ""
    ArtMeh$ = ""
Else
    h$ = RTrim(Left$(rk.txt, 33))
    Call OemToChar(h$, h$)
    l% = Len(h$)
    xAktiv% = False
    For j% = l% To 2 Step -1
        h2$ = Mid$(h$, j%, 1)
        If (h2$ = " ") Then Exit For
        If (xAktiv%) And (InStr("0123456789", h2$) <= 0) Then Exit For
        If (InStr("xX", h2$) > 0) Then xAktiv% = True
    Next j%
    If (j% > 2) Then
        ArtName$ = Left$(h$, j%)
        ArtMenge$ = Mid$(h$, j% + 1)
    Else
        ArtName$ = h$
        ArtMenge$ = ""
    End If
    ArtMeh$ = Mid$(rk.txt, 34)
End If

EK# = rk.aep

KontrollChar$ = " "
If (ProgrammChar$ <> "B") And (rk.Lief > 0) Then
    KontrollChar$ = Chr$(214)
End If

With flxarbeit(0)
    col% = .col

    iItalic% = False
    lBack& = vbButtonFace

    If (rk.loesch) Then
        lBack& = wpara.FarbeDunklerBereich
        iItalic% = True
    End If
    
    If (rk.aktivlief > 0) Then
        lBack& = vbGreen
    End If
    
    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellFontItalic = iItalic%
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle

    .col = 1
    .CellFontName = "Symbol"
    If (FirstRed%) Then
        .CellBackColor = vbRed
    End If

    row% = .row
    .TextMatrix(row%, 0) = rk.pzn
    .TextMatrix(row%, 1) = KontrollChar$
    .TextMatrix(row%, 2) = ArtName$
    .TextMatrix(row%, 3) = ArtMenge$
    .TextMatrix(row%, 4) = ArtMeh$
    .TextMatrix(row%, 5) = Abs(rk.bm)
    .TextMatrix(row%, 6) = Mid$(rk.WuAblDatum, 3)
    .TextMatrix(row%, 7) = LiefName$
    .TextMatrix(row%, 8) = SRT$
    .TextMatrix(row%, 9) = rk.BekLaufNr
    
    .col = col%
End With

Call DefErrPop
End Sub

Sub NeuerRueckKauf(RueckKaufZusatz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NeuerRueckKauf")
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
Dim i%, km%, rkMax%, erg%, x1%, xheute%, ind%
Dim aep#, AVP#
Dim txt$, xpzn$, AbCode$, wg$, X$, h$

If (ManuellPzn$ = "9999999") Then
    txt$ = ManuellTxt$
    Call CharToOem(txt$, txt$)
    wg$ = "9"
    AbCode$ = "A"
    aep# = 0#
    AVP# = 0#
Else
    txt$ = ast.kurz + ast.meng + ast.meh
    
    wg$ = Left$(ast.wg, 1)
    AbCode$ = Mid$(" AX", InStr("A ", ast.abl) + 1, 1)
            
    aep# = ast.aep
    AVP# = ast.AVP
            
    If (wg$ = "3") Then
    '    If val(ast.PZN2$) = 0 Or val(ast.Meng2$) = 0 Then GoSub Einwieger
    '    Mid$(txt$, 29, 5) = ast.Meng2$
    '    If val(ast.PZN2$) = 0 Then apzn$ = "9999999"
    End If
End If

        
X$ = txt$
X$ = LTrim$(X$)
X$ = RTrim$(X$)
If (ManuellPzn$ = "9999999") Then
    For i% = 1 To Len(X$)
        Mid$(X$, i%, 1) = UCase$(Mid$(X$, i%, 1))
    Next i%
End If
txt$ = Left$(X$ + Space$(35), 35)
      

km% = Abs(ManuellBm%)
rk.pzn = ManuellPzn$
rk.txt = txt$
rk.Lief = 0
'If (ManuellSsatz% > 0) Then ww.lief = ass.lief

rk.bm = ManuellBm%
rk.asatz = ManuellAsatz%
rk.ssatz = ManuellSsatz%
rk.best = " "
rk.nm = ManuellNm%
rk.aep = aep#
rk.abl = AbCode$
rk.wg = wg$
rk.AVP = AVP#
rk.km = km%
rk.absage = 0
rk.angebot = 0
rk.auto = Chr$(0)

rk.alt = Chr$(0)
If (ManuellSsatz% > 0) Then
    x1% = ass.lld
    h$ = Format(Now, "DDMMYY")
    xheute% = iDate(h$)
    If ((x1% + para.MonNBest) < xheute%) Then rk.alt = "?"
End If


rk.nnart = 0
rk.NNAEP = 0#
rk.besorger = " "
rk.AbholNr = 0

rk.aktivlief = 0
rk.aktivind = 0

'ww.beklaufnr = Val(Format(Day(Date), "00") + Right$(Format(Now, "HHMMSS"), 4) + Right$(ManuellPzn$, 3))
rk.BekLaufNr = CalcLaufNr&(rk.pzn)
AnzeigeLaufNr& = rk.BekLaufNr
    

rk.zugeordnet = Chr$(0)
rk.zukontrollieren = Chr$(0)
rk.fixiert = Chr$(0)
rk.DirektTyp = 0

For i% = 0 To 5
    rk.actkontrolle(i%) = 111
Next i%
rk.actzuordnung = 111
rk.PosLag = 0

rk.herst = ast.herst

'If (ast.Rez = "SG") Then
'    ww.IstBtm = 1
'Else
'    ww.IstBtm = 0
'End If

rk.loesch = 0
rk.status = 1
      
rk.IstAltLast = 0
rk.WuStatus = 0
rk.LmStatus = 0
rk.RmStatus = 0
rk.LmAnzGebucht = 0
rk.RmAnzGebucht = 0

rk.IstSchwellArtikel = 0
rk.OrgZeit = 0
                            
rk.WuAblDatum = Space$(6)

    
    
ind% = InStr(RueckKaufZusatz$, "@")
h$ = Left$(RueckKaufZusatz$, ind% - 1)
RueckKaufZusatz$ = Mid$(RueckKaufZusatz$, ind% + 1)
rk.bm = Val(h$)
rk.nm = 0

ind% = InStr(RueckKaufZusatz$, "@")
h$ = Left$(RueckKaufZusatz$, ind% - 1)
RueckKaufZusatz$ = Mid$(RueckKaufZusatz$, ind% + 1)
If (Trim$(h$) <> "") Then rk.WuAblDatum = "01" + h$




rk.GetRecord (1)
rkMax% = rk.erstmax
rkMax% = rkMax% + 1
rk.erstmax = rkMax%
rk.PutRecord (1)
rk.PutRecord (rkMax% + 1)

Call DefErrPop
End Sub

Sub mnuBearbeitenZusatz_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBearbeitenZusatz_Click")
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
Dim i%, erg%, Max%, LetztArtikelListe%
Dim pzn$, h$

tmrAction.Enabled = False

Select Case Index
    
    Case 0
        If (iMsgBox("Blindbestellung durchführen ?", vbYesNo Or vbDefaultButton2) = vbYes) Then
            Call AuslesenRueckKauf2
            If (AnzBestellArtikel% > 0) Then
                BlindBestellung% = True
                Call HoleLieferantenDaten(RueckKaufLieferant%)
                Call SucheSendeArtikel
                Call SaetzeVorbereiten
                Call SaetzeSenden
                Call UpdateRueckKaufDat(RueckKaufLieferant%, True)
                BlindBestellung% = False
                Call AuslesenRueckKauf
            End If
        End If
    
End Select

tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub AuslesenRueckKauf2()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenRueckKauf2")
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
Dim i%, j%, Lief%, Max%, l%, AltLief%, ind%, row%, RkCounter%
Dim Preis#, ZeilenWert#, EK#
Dim h$, h2$, SRT$, autox$, AktPzn$, tx$
Dim Lac$
Static IstAktiv%

If (IstAktiv% = True) Then Call DefErrPop: Exit Sub

IstAktiv% = True

rk.GetRecord (1)

Max% = rk.erstmax
RkCounter% = rk.erstcounter

frmAction!lstSortierung.Clear

AnzBestellArtikel% = 0

For i% = 1 To Max%
    rk.GetRecord (i% + 1)

    If (rk.status = 1) And (rk.loesch = 0) And (rk.aktivlief = 0) Then
        AnzBestellArtikel% = AnzBestellArtikel% + 1
        h$ = Left$(rk.txt, 18) + Mid$(rk.txt, 29) + Format(i%, "0000") + rk.pzn + Format(Abs(rk.bm), "0000") + Format(rk.BekLaufNr, String(9, 48))
        frmAction!lstSortierung.AddItem h$
                
        rk.aktivlief = RueckKaufLieferant%
        If (Wbestk2ManuellSenden%) Then
            rk.aktivind = -AnzBestellArtikel%
        Else
            rk.aktivind = AnzBestellArtikel%
        End If
        
        rk.PutRecord (i% + 1)
    End If
Next i%

If (AnzBestellArtikel% > 0) Then
    RkCounter% = (RkCounter% + 1) Mod 100
    rk.erstcounter = RkCounter%
    rk.PutRecord (1)
End If

IstAktiv% = False

Call DefErrPop
End Sub

Sub ToggleRueckKaufZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToggleRueckKaufZeile")
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
Dim i%, rInd%

tmrAction.Enabled = False

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
    If (rk.Lief > 0) Then
        rk.Lief = 0
        rk.status = 1
    Else
        rk.Lief = rk.lief1
        rk.status = 2
    End If
    rk.PutRecord (rInd% + 1)
    
    flxarbeit(0).redraw = False
    Call ZeigeRueckKaufZeile(rInd%)
    Call ErhoeheRueckKaufCounter
    flxarbeit(0).redraw = True
End If

tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub ErhoeheRueckKaufCounter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErhoeheRueckKaufCounter")
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
    
rk.GetRecord (1)
rk.erstcounter = (rk.erstcounter + 1) Mod 100
rk.PutRecord (1)

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
        Call ToggleRueckKaufZeile
        Call NaechsteRueckKaufZeile
    ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
        Call SelectZeile(UCase(Chr$(KeyAscii)))
    End If
End If

Call DefErrPop
End Sub

Sub NaechsteRueckKaufZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NaechsteRueckKaufZeile")
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

With flxarbeit(0)
    If (.row < .Rows - 1) Then
        .row = .row + 1
        If (.TopRow + opBereich.ArbeitAnzZeilen - 1 <= .row) Then
            .TopRow = .row
        End If
    End If
End With

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

Call PruefeObUpdate

Call DefErrPop
End Sub

Sub PruefeObUpdate()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeObUpdate")
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

ret% = False

tmrAction.Enabled = False

rk.GetRecord (1)
If (GlobRkMax% <> rk.erstmax) Or (RkCounter% <> rk.erstcounter) Then
    ret% = True
End If

Call opToolbar.UpdateFlag(ret%)

tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub LoescheRueckKaufZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheRueckKaufZeile")
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
Dim rInd%, Geloescht%
Dim pzn$, ch$

tmrAction.Enabled = False

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then

    Geloescht% = rk.loesch
    If (Geloescht%) Then
        rk.loesch = 0
    Else
        rk.loesch = 1
    End If
    
    rk.PutRecord (rInd% + 1)
    Call ZeigeRueckKaufZeile(rInd%)
    Call ErhoeheRueckKaufCounter
End If

tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub HighlightZeile(Optional NurNormalMachen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HighlightZeile")
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
Dim aRow%, aCol%, rInd%, ZeilenWechsel%
Dim RkLaufNr&
Dim h$, KalkText$, DirektWerte$
Static aRkLaufNr&, aFlexRow%

ZeilenWechsel% = False

With flxarbeit(0)
    If (NurNormalMachen%) Then
        .HighLight = flexHighlightNever
        KeinRowColChange% = True
        
        aRow% = .row
        aCol% = .col
        
        If (aFlexRow% < .Rows) Then
            .FillStyle = flexFillRepeat
            .row = aFlexRow%
            .col = 0
            .ColSel = .Cols - 1
            .CellForeColor = .ForeColor
            .FillStyle = flexFillSingle
            .col = aCol%
        End If
        
        aRkLaufNr& = -1&
        aFlexRow% = .row
        .HighLight = flexHighlightWithFocus
        KeinRowColChange% = False
    Else
        RkLaufNr& = Val(.TextMatrix(.row, 9))
        If (RkLaufNr& <> aRkLaufNr&) Or (aFlexRow% <> .row) Then
            
            .HighLight = flexHighlightNever
            KeinRowColChange% = True
        
            aRow% = .row
            aCol% = .col
            
            .FillStyle = flexFillRepeat
            
            If (aFlexRow% < .Rows) Then
                .row = aFlexRow%
                .col = 0
                .ColSel = .Cols - 1
                .CellForeColor = .ForeColor
                .row = aRow%
            End If
            
            .col = 0
            .ColSel = .Cols - 1
            
'            If (wpara.FarbeAktZeile) Then
'                .CellForeColor = vbHighlight
'            Else
'                .CellForeColor = vbHighlightText
'            End If
            .CellForeColor = wpara.FarbeAktZeile
            
            .FillStyle = flexFillSingle
            .col = aCol%
                       
            Call FormKurzInfo
            aRkLaufNr& = RkLaufNr&
            aFlexRow% = .row
            .HighLight = flexHighlightWithFocus
            KeinRowColChange% = False
            
            ZeilenWechsel% = True
        End If
        
    End If
End With

Call DefErrPop
End Sub

Sub RkFertigAlle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RkFertigAlle")
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
Dim i%, j%, Lm%, BuchM%, rInd%, AltRow%, AltCol%, found%, TuEs%, gef%, LoescheRowaLs%
Dim AlleBearbeitet%, SollPreisKalk%, WirdAltLast%, WumsatzMenge%, iRabatt%, bkMax%
Dim OrgGesamtWert#, Preis#, ZeilenWert#
Dim WumsatzDatum$

With flxarbeit(0)
    
'    If (EinzelneRk% = False) Then
'        Call iMsgBox("Aktion nicht durchführbar, da noch keine Rückkauf-Anfrage ausgewählt wurde !", vbInformation)
'        Call DefErrPop: Exit Sub
'    End If
    
    .redraw = False
    
    ww.SatzLock (1)
    For i% = 1 To (.Rows - 1)
        .row = i%
        
        If (Trim(.TextMatrix(i%, 1)) <> "") Then
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                
                rk.WuNeuLm = -rk.bm
                rk.WuNeuRm = -rk.bm
            
                rk.WuRm = rk.bm
                rk.WuLm = rk.bm
                
                rk.bm = 0
                rk.nm = 0
                rk.WuBm = 0
                rk.WuNm = 0
                
                rk.IstAltLast = 1
                
                ww.RawData = rk.RawData

                ww.GetRecord (1)
                bkMax% = ww.erstmax
                bkMax% = bkMax% + 1
                ww.erstmax = bkMax%
                ww.PutRecord (1)
                ww.PutRecord (bkMax% + 1)
                
                rk.status = 0
                rk.loesch = 1
                rk.PutRecord (rInd% + 1)
            End If
        End If
    Next i%
    ww.SatzUnLock (1)
        
    Call EntferneGeloeschteZeilen(0)
        
    
    .redraw = True
    KeinRowColChange% = False
    
    RkLifDat$ = Chr$(0)
    Call AuslesenRueckKauf
    
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

Call opToolbar.SpeicherIniToolbar
Set opToolbar = Nothing
Set InfoMain = Nothing

Call DefErrPop
End Sub

Sub DruckeRueckKauf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeRueckKauf")
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
Dim i%, j%, pos%, sp%(9), SN%, Y%, Max%, ind%, DRUCKHANDLE%, iLief%, iRufzeit%, abl%, anz%, aRow%, rInd%
Dim EK#, SendWert#
Dim header$, tx$, h$, KopfZeile$, LiefName$

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")

frmAction.lstSortierung.Clear

With flxarbeit(0)
    Max% = .Rows - 1
    aRow% = .row
    .redraw = False
    
    For i% = 1 To Max%
        .row = i%
        rInd% = SucheFlexZeile(True)
    
        If (rInd% > 0) Then
            If (rk.loesch = 0) And (rk.aktivlief = 0) Then
    
                EK# = rk.aep
        
                                
                h$ = Left$(rk.txt, 28) + vbTab + Mid$(rk.txt, 29, 5)
                h$ = h$ + vbTab + Mid$(rk.txt, 34, 2)
        
                tx$ = Format(Abs(rk.bm), "0")
                h$ = h$ + vbTab + tx$
                
                FabsErrf% = ass.IndexSearch(0, rk.pzn, FabsRecno&)
                If (FabsErrf% = 0) Then
                    ass.GetRecord (FabsRecno& + 1)
                
                    tx$ = ""
                    If (Abs(ass.lldat%(4) > 0)) Then
                        tx$ = sDate(ass.lldat%(4))
        '                        ret$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
                    End If
                    h$ = h$ + vbTab + tx$
                
                    tx$ = ""
                    If (Abs(ass.llrm%(4) > 0)) Then
                        tx$ = Format(ass.llrm%(4), "0")
                    End If
                    h$ = h$ + vbTab + tx$
                    
                    tx$ = ""
                    iLief% = ass.llief%(4)
                    If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
                        LiefName$ = ""
                        lif.GetRecord (iLief% + 1)
                        tx$ = lif.kurz
                        If (Trim(tx$) = "") Or (Asc(Left$(tx$, 1)) < 32) Then tx$ = Format(iLief%, "0")
                    End If
                    h$ = h$ + vbTab + tx$
                
                    tx$ = ""
                    If (ass.PosLag > 0) Then
                        abl% = ass.abl2
                        If (ass.abl1 > abl%) Then abl% = ass.abl1
                        If (abl% > 0) Then
                            tx$ = Mid$(sDate(abl%), 3)
                        End If
                    End If
                    h$ = h$ + vbTab + tx$
                    
                    tx$ = Format(ass.PosLag, "0")
                    h$ = h$ + vbTab + tx$
                Else
                    tx$ = ""
                    For j% = 0 To 4
                        h$ = h$ + vbTab + tx$
                    Next j%
                End If
                
                tx$ = Format(EK#, "0.00")
                h$ = h$ + vbTab + tx$
                
                tx$ = Format(EK# * Abs(rk.bm), "0.00")
                h$ = h$ + vbTab + tx$
                
                tx$ = Format(EK# * Abs(rk.bm), "0.00")
                h$ = h$ + vbTab + tx$
                
                tx$ = Format(Abs(rk.bm), "0")
                h$ = h$ + vbTab + tx$
                
                h$ = h$ + Chr$(10) + rk.pzn
                frmAction!lstSortierung.AddItem h$
            End If
        End If
    Next i%
    
    .row = aRow%
    .redraw = True
End With

'''''''''''''
AnzDruckSpalten% = 12
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .Titel = "P Z N"
    .TypStr = String$(7, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(1)
    .Titel = "A R T I K E L"
    .TypStr = String$(25, "X")  '28
    .Ausrichtung = "L"
End With
With DruckSpalte(2)
    .Titel = ""
    .TypStr = String$(6, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(3)
    .Titel = ""
    .TypStr = String$(4, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(4)
    .Titel = "B M"
    .TypStr = String$(4, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(5)
    .Titel = "L.Dat."
    .TypStr = String$(6, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(6)
    .Titel = "L M"
    .TypStr = String$(3, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(7)
    .Titel = "L.Lief."
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(8)
    .Titel = "Verfall"
    .TypStr = String$(5, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(9)
    .Titel = "Lager"
    .TypStr = String$(5, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(10)
    .Titel = "A E P"
    .TypStr = "99999.99"
    .Ausrichtung = "R"
End With
With DruckSpalte(11)
    .Titel = "Wert"
    .TypStr = "99999.99"
    .Ausrichtung = "R"
End With

Call InitDruckZeile(True)
        

DruckSeite% = 0
AnzFaxDruckArtikel% = 0
AnzFaxDruckPackungen% = 0
FaxDruckWert# = 0#
FaxDruckWert2# = 0#
Call RueckKaufDruckKopf
            
With frmAction!lstSortierung
    anz% = .ListCount
    For i% = 1 To anz%
        .ListIndex = i% - 1
        h$ = .text + vbTab
        
        ind% = InStr(h$, Chr$(10))
        If (ind% > 0) Then
            h$ = Mid$(h$, ind% + 1) + Left$(h$, ind% - 1) + vbTab
        End If
                    
        Call FaxDruckZeile(h$)
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1200) Then
            Call DruckFuss
            Call RueckKaufDruckKopf
        End If
    Next i%
    
    Call RueckKaufDruckSumme
    Call DruckFuss(False)
    
    Printer.EndDoc
End With

Call StopAnimation(frmAction)

Call DefErrPop
End Sub

Sub RueckKaufDruckKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RueckKaufDruckKopf")
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
Dim l&, he&, i%, pos%, X%, Y%, GesBreite&
Dim h$, heute$, SeitenNr$, header$, LinksOben$

header$ = "?"
If (ProgrammChar$ = "B") Then
    If (RueckKaufLieferant% > 0) Then
        lif.GetRecord (RueckKaufLieferant% + 1)
        header$ = RTrim$(lif.kurz) + "  (" + Mid$(Str$(RueckKaufLieferant%), 2) + ")"
    End If
    LinksOben$ = "Rückkauf-Anfrage"
Else
    header$ = lblArbeit(0).Caption
    LinksOben$ = "gesendete Rückkauf-Anfrage"
End If

heute$ = Format(Day(Date), "00") + "-"
heute$ = heute$ + Format(Month(Date), "00") + "-"
heute$ = heute$ + Format(Year(Date), "0000")
heute$ = heute$ + " " + Left$(Time$, 5)

With Printer
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 14
    Printer.Print LinksOben$
    
    DruckSeite% = DruckSeite% + 1
    SeitenNr$ = "-" + Str$(DruckSeite%) + " -"
    l& = .TextWidth(SeitenNr$)
    .CurrentX = (.ScaleWidth - l&) / 2: .CurrentY = 0
    Printer.Print SeitenNr$
    
    l& = .TextWidth(heute$)
    .CurrentX = .ScaleWidth - l& - 10: .CurrentY = 0
    Printer.Print heute$
    
    .CurrentX = 0
    l& = .TextHeight("A") + 500
    .CurrentY = l&
    
    .Font.Size = 18
    l& = .TextWidth(header$)
    he& = .TextHeight("A")
        
    X% = (.ScaleWidth - (l& + 800)) / 2
    Y% = .CurrentY
    .DrawWidth = 2
    RoundRect .hdc, X% / .TwipsPerPixelX, Y% / .TwipsPerPixelY, (X% + l& + 800) / .TwipsPerPixelX, (Y% + he& + 800) / .TwipsPerPixelY, 200, 200
    .CurrentX = (.ScaleWidth - l&) / 2
    .CurrentY = Y% + 400
    Printer.Print header$
    
    Printer.Print
    .Font.Size = DruckFontSize%
    Printer.Print
    Printer.Print
    
    For i% = 0 To (AnzDruckSpalten% - 1)
        h$ = RTrim(DruckSpalte(i%).Titel)
        If (DruckSpalte(i%).Ausrichtung = "L") Then
            X% = DruckSpalte(i%).StartX
        Else
            X% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h$)
        End If
        .CurrentX = X%
        Printer.Print h$;
    Next i%
    
    Printer.Print " "
    
    Y% = Printer.CurrentY
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
End With

Call DefErrPop
End Sub

Sub RueckKaufDruckSumme()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RueckKaufDruckSumme")
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
Dim Y%, GesBreite&, tx$
            
Y% = Printer.CurrentY
GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

Y% = Printer.CurrentY
Printer.CurrentY = Y% + 30

Printer.CurrentX = DruckSpalte(1).StartX
Printer.Print Format(AnzFaxDruckArtikel%, "0") + " Position(en) / " + Format(AnzFaxDruckPackungen%, "0") + " Packung(en)";
tx$ = Format(FaxDruckWert#, "0.00")     'MarkWert#
Printer.CurrentX = GesBreite& - Printer.TextWidth(tx$)
Printer.Print tx$;
Printer.Print " "

Call DefErrPop
End Sub

