VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAction 
   Caption         =   "Personal-Verkaufstatistik"
   ClientHeight    =   7890
   ClientLeft      =   390
   ClientTop       =   630
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
   Icon            =   "WinPvs.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   7890
   ScaleWidth      =   11295
   Begin VB.PictureBox picAusdruck 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDatei 
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   14
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
   Begin VB.PictureBox picAnimationBack 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   4080
      ScaleHeight     =   2370
      ScaleWidth      =   5625
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   5655
      Begin ComCtl2.Animation aniAnimation 
         Height          =   1095
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1931
         _Version        =   327681
         Center          =   -1  'True
         BackColor       =   -2147483624
         FullWidth       =   73
         FullHeight      =   73
      End
      Begin VB.Label lblAnimation 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         Caption         =   "Aufgabe wird bearbeitet ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   180
         Width           =   5355
      End
   End
   Begin VB.ListBox lstSortierung 
      Height          =   300
      Left            =   7800
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Begin VB.ComboBox cboMitarbeiter 
         Height          =   360
         Left            =   1560
         Style           =   2  'Dropdown-Liste
         TabIndex        =   17
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton cmdEsc 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   450
         Index           =   0
         Left            =   5280
         TabIndex        =   12
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1200
      End
      Begin VB.PictureBox picSummenzeile 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   1320
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   3960
         Index           =   0
         Left            =   240
         TabIndex        =   7
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
         HighLight       =   2
         GridLines       =   0
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   1500
         Index           =   0
         Left            =   600
         TabIndex        =   9
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
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfoZusatz 
         Height          =   780
         Index           =   0
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5280
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
      Begin VB.Label lblPrmBasis 
         Caption         =   "Prämienbasis:"
         Height          =   315
         Left            =   7560
         TabIndex        =   19
         Top             =   5700
         Width           =   2235
      End
      Begin VB.Label lblMitarbeiter 
         Caption         =   "&Mitarbeiter"
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
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1935
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
         TabIndex        =   8
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
         TabIndex        =   10
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
            Picture         =   "WinPvs.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":059C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":082E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":0B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":0C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":0EEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":1498
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":172A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":1A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":1D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":1FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":230A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":259C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":282E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":2AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":2DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":306C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":32FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":3590
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":3822
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":3B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":3E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":4170
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":448A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   0
      Left            =   10200
      Top             =   120
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
            Picture         =   "WinPvs.frx":47A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":48B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":49C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":4CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":4DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5086
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5318
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":542A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":553C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5856
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":5F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":60AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":61C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":62D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":65EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":66FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":6810
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":6922
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":6A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":6D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":7068
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":7382
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "WinPvs.frx":769C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDateiInd 
         Caption         =   "&PVS"
         Index           =   0
      End
      Begin VB.Menu mnuDummy5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuBearbeiten 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Neue Auswertung"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Daten überleiten"
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
         Caption         =   ""
         Index           =   5
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
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
         Caption         =   "&Vorschau"
         Index           =   9
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   10
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&PBA"
         Index           =   11
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Diagramme"
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
         Caption         =   "&ApoControl"
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
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Sonstige F&unktionen"
         Index           =   0
         Begin VB.Menu mnuBearbeitenZusatzInd 
            Caption         =   ""
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuAnsicht 
      Caption         =   "&Ansicht"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Symbolleiste"
         Index           =   0
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
         Index           =   0
      End
      Begin VB.Menu mnuFarbe 
         Caption         =   "&Farbeinstellungen"
         Begin VB.Menu mnuFarbeInd 
            Caption         =   "&Arbeitsbereich ..."
            Index           =   0
         End
         Begin VB.Menu mnuFarbeInd 
            Caption         =   "&Infobereich ..."
            Index           =   1
         End
         Begin VB.Menu mnuFarbeInd 
            Caption         =   "&Dunkler Bereich ..."
            Index           =   2
         End
         Begin VB.Menu mnuFarbeInd 
            Caption         =   "Aktuelle &Zeile ..."
            Index           =   3
         End
      End
      Begin VB.Menu mnuRahmenAnzeigen 
         Caption         =   "&Rahmen anzeigen"
      End
      Begin VB.Menu mnuDummy7 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Schrift von Inf&ormationen ..."
         Index           =   0
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Schrift von Te&xten ..."
         Index           =   1
      End
      Begin VB.Menu mnuDummy9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZeilen 
         Caption         =   "&leere Tabellenzeilen einblenden"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuOptionen 
         Caption         =   "&Optionen ..."
      End
      Begin VB.Menu mnuDummy4 
         Caption         =   "-"
      End
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

Const INI_SECTION = "PVS"
Const INFO_SECTION = "Infobereich PVS"


'Dim scrAuswahlAltValue%
Dim InRowColChange%

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

Dim HochfahrenAktiv%

Dim Standard%

Dim HatKunden As Boolean

Private Const DefErrModul = "WINPVS.FRM"

Public Sub ApoControl()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ApoControl")
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
Dim ApoVKID  As Integer
Dim SQLStr As String

If Not ApoCDBda Then ApoCDBda = OpenCreateControlDB
If Not ApoCDBda Then
  Call DefErrPop
  Exit Sub
End If

ApoControlTRec.MoveFirst
Do While Not ApoControlTRec.EOF
  If ApoControlTRec!Name = "Personal" Then
    ApoVKID = ApoControlTRec!Id
  End If
  ApoControlTRec.MoveNext
Loop
ApoControlRec.Close
Set ApoControlRec = ApoControlDB.OpenRecordset("SELECT * FROM BEZEICHNUNGEN WHERE TabellenID = " + CStr(ApoVKID))
Set ApoControlWRec = ApoControlWDB.OpenRecordset("Werte", dbOpenTable)

ApoControlRec.MoveFirst
Do While Not ApoControlRec.EOF
  SQLStr = "DELETE FROM WERTE WHERE BezID = " + CStr(ApoControlRec!Id) + " AND Datevalue(Datum) >= Datevalue(" + Chr(34) + Format(vonAuswD, "dd.mm.yyyy") + Chr(34) + ") AND Datevalue(Datum) <= Datevalue(" + Chr(34) + Format(bisAuswD, "dd.mm.yyyy") + Chr(34) + ")"
  ApoControlWDB.Execute (SQLStr)
  ApoControlWRec.AddNew
  ApoControlWRec!BezID = ApoControlRec!Id
  ApoControlWRec!datum = vonAuswD
  Select Case ApoControlRec!Id
  Case 81   '"Anzahl Normtage"
    ApoControlWRec!Betrag = xVal(aInfo$(0, 3, 0))
  Case 82   'Kd/NT
    ApoControlWRec!Betrag = xVal(aInfo$(1, 3, 0))
  Case 83   'Barverkauf/Barverkkd
    ApoControlWRec!Betrag = xVal(aInfo$(0, 5, 0))
  Case 84   '% Zusatzverk/Rezeptkd
    ApoControlWRec!Betrag = xVal(aInfo$(5, 3, 0))
  End Select
  ApoControlWRec.Update
  ApoControlRec.MoveNext
Loop

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
Dim i%, ind%
Dim h$

Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        mnuExtras.Enabled = True
        
        mnuBearbeitenInd(MENU_F2).Enabled = True
        mnuBearbeitenInd(MENU_F3).Enabled = True
        mnuBearbeitenInd(MENU_F4).Enabled = False
        mnuBearbeitenInd(MENU_F5).Enabled = False
        mnuBearbeitenInd(MENU_F6).Enabled = True
        mnuBearbeitenInd(MENU_F7).Enabled = False
        mnuBearbeitenInd(MENU_F8).Enabled = False
        mnuBearbeitenInd(MENU_F9).Enabled = False
        mnuBearbeitenInd(MENU_SF2).Enabled = True
        mnuBearbeitenInd(MENU_SF3).Enabled = False
        mnuBearbeitenInd(MENU_SF4).Enabled = True
        mnuBearbeitenInd(MENU_SF5).Enabled = True
        mnuBearbeitenInd(MENU_SF6).Enabled = False
        mnuBearbeitenInd(MENU_SF7).Enabled = False
        mnuBearbeitenInd(MENU_SF8).Enabled = True
        mnuBearbeitenInd(MENU_SF9).Enabled = False
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        flxarbeit(0).BackColorSel = vbHighlight
        flxInfo(0).BackColorSel = vbHighlight
End Select

For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%

ProgrammModus% = NeuerModus%

Call DefErrPop
End Sub

Private Sub ZeilenEinAus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeilenEinAus")
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
Me.MousePointer = vbHourglass
With flxarbeit(0)
  For i% = 1 To .Rows - 1
    If mnuZeilen.Checked Then
      .RowHeight(i%) = .RowHeight(0)
    Else
      If Val(.TextMatrix(i%, 4)) = 0 Then
        .RowHeight(i%) = 0
      End If
    End If
  Next i%
End With
Me.MousePointer = vbNormal
Call DefErrPop
End Sub

Private Sub cmdDatei_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdDatei_Click")
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
Dim erg%

ProgrammTyp% = Index
Call InitProgrammTyp
HochfahrenAktiv% = False
Call Form_Resize
'Call InitProgramm
flxarbeit(0).SetFocus

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

Private Sub flxarbeit_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(flxarbeit(Index), picBack(Index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub flxarbeit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_KeyDown")
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

'If (Index = 0) Then
'    If (Shift And 2) And (KeyCode = vbKeyV) Then
'        EingabeStr$ = Clipboard.GetText
'    End If
'End If

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

'If (Index = 0) Then
'    If (KeyAscii = vbKeySpace) Then
'        Call ToggleBestellZeile
'        Call NaechsteBestellZeile
'    ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
'        Call frmAction.SelectZeile(UCase(Chr$(KeyAscii)))
'    End If
'End If

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

With flxarbeit(0)
    .HighLight = flexHighlightNever
End With

'EingabeStr$ = ""
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

Private Sub flxInfo_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(flxInfo(Index), picBack(Index), Source, x, Y)
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
'EingabeStr$ = ""

With flxarbeit(0)
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

'If (KeinRowColChange% = False) Then
'    Call EchtKurzInfo
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

'If (Index = 0) Then
'    If ((flxarbeit(0).Redraw = True) And (KeinRowColChange% = False)) Then
'        Call HighlightZeile
'        flxInfo(0).row = 0
'        flxInfo(0).col = 0
'    End If
'End If
    
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


On Error Resume Next

HochfahrenAktiv% = True
   
Call wpara.InitEndSub(Me)
Set opToolbar = New clsToolbar

Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI)
Call wpara.InitFont(Me)
Call HoleIniWerte

Set InfoMain = New clsInfoBereich
Set opBereich = New clsOpBereiche


Call InitDateiButtons

Call InitAnimation

Call InitProgrammTyp

Me.SetFocus
DoEvents

HochfahrenAktiv% = False

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

HochfahrenAktiv% = True

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
    .Visible = True
End With

h$ = ProgrammNamen$(ProgrammTyp%)
Caption = h$
ProgrammChar$ = Left$(h$, 1)

On Error Resume Next
For i% = 1 To 5
    Unload mnuBearbeitenZusatzInd(i%)
Next i%
On Error GoTo DefErr

''''''''''''''''
Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION)

'.cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
cmdToolbar(1).ToolTipText = "F2 neue Auswertung"
cmdToolbar(2).ToolTipText = "F3 Daten überleiten"
'.cmdToolbar(3).ToolTipText = "F4 Quittieren der Nachricht"
'.cmdToolbar(4).ToolTipText = "F5 Entfernen eines Artikels oder einer Bedingung"
cmdToolbar(5).ToolTipText = "F6 Drucken"
'.cmdToolbar(6).ToolTipText = "F7 Aktualisieren"
'.cmdToolbar(7).ToolTipText = "F8 Zusatztext"
cmdToolbar(8).ToolTipText = "F9 Abmelden"
cmdToolbar(9).ToolTipText = "shift+F2 tageweise Vorschau ohne Überleitung"
'.cmdToolbar(10).ToolTipText = "shift+F3 akutellen Lieferanten wechseln"
cmdToolbar(11).ToolTipText = "shift+F4 PBA"
cmdToolbar(12).ToolTipText = "shift+F5 Diagramme"
'.cmdToolbar(13).ToolTipText = "shift+F6 Datenübertragung zum GH"
'.cmdToolbar(14).ToolTipText = "shift+F7 Warenwert"
cmdToolbar(15).ToolTipText = "shift+F8 Speichern für ApoControl"
'.cmdToolbar(16).ToolTipText = "shift+F9 Rückkauf-Anfrage"
''cmdToolbar(19).ToolTipText = "Programm beenden"

'Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, INFO_SECTION)
'Call InfoMain.ZeigeInfoBereich("", False)
Call ZeigeInfoBereich(False)
flxInfo(0).row = 0
flxInfo(0).col = 0

Call opBereich.InitBereich(Me, opToolbar)
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = True
opBereich.ArbeitWasDarunter = True
opBereich.InfoTitel = False
opBereich.InfoZusatz = 0
opBereich.InfoAnzZeilen = 7 'InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = 0

With flxarbeit(0)
    .Cols = 17
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 1
    .FormatString = ">Tag|>PrämBasis|>Zus.Basis|>Snd.Basis|>AnzKd|>AnzRez|>%PräKd|>%PrivKd|>%RezKd|>Erster|>Letzter|>kleine|>Pausen|>große|>Pausen|>"
End With
    
'mnuBearbeitenInd(10).Caption = "&Lieferant"
'mnuBearbeitenInd(13).Caption = "&Sendevorgang"
'mnuBearbeitenInd(14).Caption = "Wa&renwert"
'mnuBearbeitenInd(15).Caption = "Bestell&vorschlag"

'mnuBearbeitenZusatzInd(0).Caption = "Artikel zu W&Ü dazu"
'mnuBearbeitenZusatzInd(0).Enabled = True
'
'Load mnuBearbeitenZusatzInd(1)
'mnuBearbeitenZusatzInd(1).Caption = "&Blind-Bestellung"
'mnuBearbeitenZusatzInd(1).Enabled = True
    
mnuRahmenAnzeigen.Checked = wpara.FarbeRahmen


picBack(0).Visible = True

mnuZeilen.Checked = LeereZeilen

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
Call ProgrammEnde
Call DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(lblArbeit(Index), picBack(Index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub lblInfo_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(lblInfo(Index), picBack(Index), Source, x, Y)
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
Dim erg%, row%, col%, i%, OrgMitarb%
Dim l&
Dim h$, mErg$
Dim Tag As Date

Select Case Index

    Case MENU_F2
      Call DatumEin
    
    Case MENU_F3
      Call Überleitung
      If ErstesMal% = 0 Then Call DatumEin
    
    Case MENU_F6
      If Vorschau Then
        Call DruckeWindow(Me)
      Else
        If xVal(aDetail$(4, 0, aTage%)) > 0 Then
          OrgMitarb% = cboMitarbeiter.ListIndex
          frmDruck.Show vbModal
          If EditErg% Mod 100 = 1 Then
            cboMitarbeiter.ListIndex = 0
            DoEvents
            If HatKunden Then Call WinPvsAusdruck
          End If
          If EditErg% >= 100 Then
            For i% = 1 To cboMitarbeiter.ListCount - 1
              cboMitarbeiter.ListIndex = i%
              DoEvents
              If HatKunden Then Call WinPvsAusdruck
            Next i%
          End If
          cboMitarbeiter.ListIndex = OrgMitarb%
        End If
      End If
    Case MENU_SF2
      sF2Vorschau = True
      Call DatumEin
      sF2Vorschau = False
      
    Case MENU_SF4
        Call PbaInit
        
    Case MENU_SF5
      If xVal(aDetail$(4, 0, aTage%)) > 0 Then
        frmDiagramm.Show vbModal
      End If
      
    Case MENU_SF8
      If xVal(aDetail$(4, 0, aTage%)) > 0 Then
        Tag = DateAdd("m", 1, vonAuswD)
        Tag = DateAdd("d", -1, Tag)
        If Tag <> bisAuswD Then
          Call MsgBox("Der Auswertungszeitraum muss genau einen Kalendermonat umfassen. Speichern für ApoControl nicht möglich.", vbInformation Or vbOKOnly)
        Else
          If (MsgBox("Möchten Sie diese Werte für Apo-Control speichern?", vbYesNo) = vbYes) Then
            Call ApoControl
          End If
        End If
      End If
End Select

Call DefErrPop
End Sub

Private Sub mnuBearbeitenZusatzInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBearbeitenZusatzInd_Click")
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
Call ActProgram.mnuBearbeitenZusatzClick(Index)
Call DefErrPop
End Sub

Private Sub mnuDateiInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDateiInd_Click")
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
If (Index = 1) Then Index = 2
cmdDatei(Index).Value = True
Call DefErrPop
End Sub

Private Sub mnuDiagramme_Click()
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

frmDiagramm.Show 1

Call DefErrPop
End Sub

Private Sub mnuFarbeInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuFarbeInd_Click")
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
Dim erg%

erg% = wpara.EditFarbe(dlg, Index)
If (erg%) Then
    Call opBereich.ResizeWindow
End If

Call DefErrPop
End Sub

Private Sub mnuFont_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuFont_Click")
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
Dim erg%

erg% = wpara.EditFont(dlg, Index)
If (erg%) Then
    Call wpara.InitFont(frmAction)
    Call opBereich.ResizeWindow

    frmAction.flxarbeit(0).Rows = 1
    'frmAction.flxInfo(0).Clear
    Call cboMitarbeiter_Click
    'Call ActProgram.mnuFontClick
End If

Call DefErrPop
End Sub

Private Sub mnuRahmenAnzeigen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuRahmenAnzeigen_Click")
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

If (mnuRahmenAnzeigen.Checked) Then
    mnuRahmenAnzeigen.Checked = False
Else
    mnuRahmenAnzeigen.Checked = True
End If

wpara.FarbeRahmen = Abs(mnuRahmenAnzeigen.Checked)
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
Dim i%, ind%, erg%
Dim h$

'If (Shift And vbCtrlMask And (KeyCode <> 17)) Then
'    ind% = 0
'    Select Case KeyCode
'        Case vbKeyF2
'            ind% = 1
'        Case vbKeyF3
'            ind% = 2
'        Case vbKeyF4
'            ind% = 3
'        Case vbKeyF5
'            ind% = 4
'        Case vbKeyF6
'            ind% = 5
'        Case vbKeyF7
'            ind% = 6
'        Case vbKeyF8
'            ind% = 7
'        Case vbKeyF9
'            ind% = 8
'        Case vbKeyS
'            ind% = -1
''        Case vbKeyF11
''            ind% = 9
'    End Select
'    If ((Shift And vbShiftMask) And (ind% > 0)) Then
'        ind% = ind% + 8
'    End If
'    If (ind% > 0) Then
'        h$ = cmdToolbar(ind%).ToolTipText
'        picToolTip.Width = picToolTip.TextWidth(h$ + "x")
'        picToolTip.Height = picToolTip.TextHeight(h$) + 45
'        picToolTip.Left = picToolbar.Left + cmdToolbar(ind%).Left
'        picToolTip.Top = picToolbar.Top + picToolbar.Height + 60
'        picToolTip.Visible = True
'        picToolTip.Cls
'        picToolTip.CurrentX = 2 * Screen.TwipsPerPixelX
'        picToolTip.CurrentY = 0
'        picToolTip.Print h$
'        KeyCode = 0
'    ElseIf (ind% = -1) Then
'        h$ = "02000"
'        If (ProgrammTyp% = 0) Then
'            If ((BestellAnzeige% = 2) And (IstDirektLief%)) Then h$ = "02091"
'            Call Stammdaten(Format(Lieferant%, "000"), Val(h$), DirektBewertung#, ActProgram)
'            Call ActProgram.MenuBearbeiten(MENU_F7)
'    '        Call ActProgram.AuslesenPalette
'        Else
'            With flxarbeit(0)
'                Call Stammdaten(Format(Asc(Left$(.TextMatrix(.row, 16), 1)), "000"), Val(h$))
'            End With
'        End If
'        KeyCode = 0
'    End If
'End If

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

frmAbout.Show vbModal

Call DefErrPop
End Sub

Private Sub mnuOptionen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuOptionen_Click")
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

frmWinPvsOptionen.Show 1

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

Private Sub mnuZeilen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuZeilen_Click")
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
Dim key$, h$

mnuZeilen.Checked = Not (mnuZeilen.Checked)
LeereZeilen = (mnuZeilen.Checked)
If LeereZeilen Then
  h$ = "J"
Else
  h$ = "N"
End If
key$ = "LeereZeilen"
l& = WritePrivateProfileString("PVS", key$, h$, CurDir + "\winop.ini")

Call ZeilenEinAus
Call DefErrPop
End Sub

Private Sub picBack_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(picBack(Index), picBack(Index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

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
'If (picToolTip.Visible = True) Then
'    picToolTip.Visible = False
'End If

Call DefErrPop
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
opToolbar.DragX = x
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

Call DefErrPop
End Sub

Public Sub EchtKurzInfo()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EchtKurzInfo")
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
'Call ActProgram.EchtKurzInfo
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
'Dim aRow%, aCol%, rInd%, ZeilenWechsel%
'Dim KalkAvp#, RundAvp#
'Dim BekLaufNr&
'Dim h$, KalkText$, DirektWerte$
'Static aBekLaufNr&, aFlexRow%, aDirektWerte$
'
'ZeilenWechsel% = False
'
'With flxarbeit(0)
'    If (NurNormalMachen%) Then
'        .HighLight = flexHighlightNever
'        KeinRowColChange% = True
'
'        aRow% = .row
'        aCol% = .col
'
'        If (aFlexRow% < .Rows) Then
'            .FillStyle = flexFillRepeat
'            .row = aFlexRow%
'            .col = 0
'            .ColSel = .Cols - 1
'            .CellForeColor = .ForeColor
'            .FillStyle = flexFillSingle
'            .col = aCol%
'        End If
'
'        aBekLaufNr& = -1&
'        aFlexRow% = .row
'        aDirektWerte$ = ""
'        .HighLight = flexHighlightWithFocus
'        KeinRowColChange% = False
'    Else
'        If (ProgrammChar$ = "B") Then
'            BekLaufNr& = Val(.TextMatrix(.row, 21))
'        Else
'            BekLaufNr& = Val(.TextMatrix(.row, 20))
'        End If
'        If (BekLaufNr& <> aBekLaufNr&) Or (aFlexRow% <> .row) Or ((IstDirektLief%) And (.col = 8)) Then
'
'            .HighLight = flexHighlightNever
'            KeinRowColChange% = True
'
'            aRow% = .row
'            aCol% = .col
'
'            .FillStyle = flexFillRepeat
'
'            If (aFlexRow% < .Rows) Then
'                .row = aFlexRow%
'                .col = 0
'                .ColSel = .Cols - 1
'                .CellForeColor = .ForeColor
'                .row = aRow%
'            End If
'
'            .col = 0
'            .ColSel = .Cols - 1
'
'            .CellForeColor = wpara.FarbeAktZeile
'
'            .FillStyle = flexFillSingle
'            .col = aCol%
'
'            Call EchtKurzInfo
'            aBekLaufNr& = BekLaufNr&
'            aFlexRow% = .row
'            aDirektWerte$ = Trim$(.TextMatrix(.row, 5)) + vbTab + Trim$(.TextMatrix(.row, 6)) + vbTab + Trim$(.TextMatrix(.row, 7)) + vbTab + Trim$(.TextMatrix(.row, 26))
'            .HighLight = flexHighlightWithFocus
'            KeinRowColChange% = False
'
'            ZeilenWechsel% = True
'        End If
'
''        If (.col = 9) And (Left$(.TextMatrix(.row, 9), 6) = "Absage") Then
'        If (ProgrammChar$ = "B") Then
'            If (.col = 8) Then
'                h$ = ToolTipSchwellwert$(Val(.TextMatrix(.row, 27)))
'                If (h$ <> "") Then
'                    picQuittieren.Visible = False
'                    Call AnzeigeKommentar(h$, 7, 2)
'                End If
'            ElseIf (.col = 10) Then
'                h$ = ToolTipAbsagen$(.TextMatrix(.row, 0))
'                If (h$ <> "") Then
'                    picQuittieren.Visible = False
'                    Call AnzeigeKommentar(h$, 9, 1)
'                End If
'            ElseIf (mnuDatei.Enabled) And (ZeilenWechsel% = False) Then
'                picQuittieren.Visible = False
'            End If
'        ElseIf (ProgrammChar$ = "W") Then
'            picQuittieren.Visible = False
'
'            If (RTrim$(.TextMatrix(.row, 0)) = "XXXXXXX") Then Call DefErrPop: Exit Sub
'
'            rInd% = SucheFlexZeile(True)
'            If (rInd% > 0) Then
'                h$ = ""
'                If (.TextMatrix(.row, 1) = "$") Then
'                    h$ = ActProgram.PruefeAutomaticPreis(RundAvp#)
'                End If
'                If (h$ <> "") Then
'                    Call AnzeigeKommentar(h$, 2, 2)
'                Else
'                    h$ = Trim(ww.WuText)
'                    If (h$ <> "") Then Call AnzeigeKommentar(h$, 2, 2)
'                End If
'            End If
'        End If
'    End If
'End With

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
'Dim i%, gef%
'Dim ch$
'
'gef% = -1
'With flxarbeit(0)
'    For i% = (.row + 1) To (.Rows - 1)
'        ch$ = Left$(.TextMatrix(i%, 2), 1)
'        If (ch$ = SearchLetter$) Then
'            gef% = i%
'            Exit For
'        End If
'    Next i%
'
'    If (gef% < 0) Then
'        For i% = 1 To (.row - 1)
'            ch$ = Left$(.TextMatrix(i%, 2), 1)
'            If (ch$ = SearchLetter$) Then
'                gef% = i%
'                Exit For
'            End If
'        Next i%
'    End If
'
'    If (gef% > 0) Then
'        Call HighlightZeile(True)
'        .row = gef%
'        .col = 8
'
'        If (.row < .TopRow) Then
'            .TopRow = .row
'        Else
'            If (.row >= (.TopRow + opBereich.ArbeitAnzZeilen - 2)) Then
'                .TopRow = .row - opBereich.ArbeitAnzZeilen + 2
'            End If
'    '        While ((.row - .TopRow) >= (ParentBereich.ArbeitAnzZeilen - 1))
'    '            .TopRow = .TopRow + 1
'    '        Wend
'        End If
'        Call HighlightZeile
'        Call EchtKurzInfo
'    End If
'End With

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
    Me.Font.Bold = True
    '.FormatString = ">Tag|>PrämBasis|>ZusatzBasis|>SonderBasis|>AnzKd|>AnzRez|>%PräKd|>%PrivKd|>%RezKd|>Erster|>Letzter|>kl.|>Pausen|>gr.|>Pausen"
    
    .ColWidth(0) = Me.TextWidth("999")
    .ColWidth(1) = Me.TextWidth("999999999")
    .ColWidth(2) = Me.TextWidth("999999999")
    .ColWidth(3) = Me.TextWidth("99999999")
    .ColWidth(4) = Me.TextWidth("99999")
    .ColWidth(5) = Me.TextWidth("999999")
    .ColWidth(6) = Me.TextWidth("9999999")
    .ColWidth(7) = Me.TextWidth("9999999")
    .ColWidth(8) = Me.TextWidth("9999999")
    .ColWidth(9) = Me.TextWidth("999:99")
    .ColWidth(10) = Me.TextWidth("999:99")
    .ColWidth(11) = Me.TextWidth("99999")
    .ColWidth(12) = Me.TextWidth("999999")
    .ColWidth(13) = Me.TextWidth("99999")
    .ColWidth(14) = Me.TextWidth("999999")
    .ColWidth(15) = 0
    .ColWidth(16) = wpara.FrmScrollHeight '+ 2 * wpara.FrmBorderHeight
    Me.Font.Bold = False
    
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + Me.TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(15) = .Width - spBreite%
End With

With flxInfo(0)
'    sp& = .Width / 8
'    .ColWidth(0) = 2 * sp&
    sp& = .Width / 3 - 30&
    .ColWidth(0) = 0
    For i% = 1 To 5 Step 2
        .ColWidth(i%) = sp& * 0.7
        .ColWidth(i% + 1) = sp& - .ColWidth(i%)
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

picSummenzeile.Font.Name = flxarbeit(0).Font.Name
picSummenzeile.Font.Size = flxarbeit(0).Font.Size
picSummenzeile.Top = flxarbeit(0).Top + flxarbeit(0).Height
picSummenzeile.Left = flxarbeit(0).Left
picSummenzeile.Height = opBereich.ZeilenHoeheY + 90
picSummenzeile.Width = flxarbeit(0).Width
picSummenzeile.Visible = True
'    Call ActProgram.ZeigeWerte
    
lblMitarbeiter.Left = wpara.LinksX
lblMitarbeiter.Top = wpara.TitelY   'FlexY%
cboMitarbeiter.Top = wpara.TitelY + (lblMitarbeiter.Height - cboMitarbeiter.Height) \ 2
cboMitarbeiter.Left = lblMitarbeiter.Left + lblMitarbeiter.Width + 150

lblPrmBasis.Top = wpara.TitelY
lblPrmBasis.Left = cboMitarbeiter.Left + cboMitarbeiter.Width + wpara.LinksX
lblPrmBasis.Width = frmAction.ScaleWidth - lblPrmBasis.Left - wpara.LinksX
  
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

lblMitarbeiter.BackColor = wpara.FarbeArbeit
lblPrmBasis.BackColor = wpara.FarbeArbeit
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

On Error Resume Next
Call opToolbar.SpeicherIniToolbar
If (WindowState = vbMaximized) Then
    l& = WritePrivateProfileString(UserSection$, "StartX", Str$(-9999), INI_DATEI)
Else
    l& = WritePrivateProfileString(UserSection$, "StartX", Str$(Left), INI_DATEI)
    l& = WritePrivateProfileString(UserSection$, "StartY", Str$(Top), INI_DATEI)
    l& = WritePrivateProfileString(UserSection$, "BreiteX", Str$(Width), INI_DATEI)
    l& = WritePrivateProfileString(UserSection$, "HoeheY", Str$(Height), INI_DATEI)
End If

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
Dim l&, f&, lVal&
Dim h$, h2$, key$
    
With frmAction
    
    .Height = wpara.WorkAreaHeight
    If (wpara.BildFaktor = 0.8!) Then
        .Width = 10980 * wpara.BildFaktor '10200
    Else
        .Width = .Width * wpara.BildFaktor
    End If

    
    
    iVal% = (Screen.Width - Width) / 2
    If (iVal% < 0) Then
        iVal% = 0
    End If
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "StartX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    lVal& = Val(h$)
    If (lVal& > 30000) Then
        iVal% = -9999
    Else
        iVal% = lVal&
    End If
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
    
    ProgrammTyp% = 0
    If (Command <> "") Then ProgrammTyp% = Val(Command)
    
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
Dim l&
Dim h$

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

Call DefErrPop
End Sub

Sub InitAnimation()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitAnimation")
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
With lblAnimation
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = TextWidth("Parameter werden eingelesen ...") + 300
    .Height = TextHeight("Äg") + 150
End With

With aniAnimation
    .Left = lblAnimation.Left + (lblAnimation.Width - .Width) / 2
    .Top = lblAnimation.Top + lblAnimation.Height + 90
End With

With picAnimationBack
    .Width = lblAnimation.Width + 2 * wpara.LinksX
    .Height = aniAnimation.Top + aniAnimation.Height + 90
End With

Call DefErrPop
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

Sub InitDateiButtons()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitDateiButtons")
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

For i% = 1 To 2
    Load cmdDatei(i%)
Next i%

For i% = 0 To 2
    cmdDatei(i%).Top = 0
    cmdDatei(i%).Left = i% * 900
    cmdDatei(i%).Visible = True
    cmdDatei(i%).ZOrder 1
Next i%

cmdDatei(0).Caption = "&L"
cmdDatei(1).Caption = "&K"
cmdDatei(2).Caption = "&W"

Call DefErrPop
End Sub

Sub ZeigeInfoBereich(AuchWerte%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInfoBereich")
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
Dim h$, txWert$, tx$

With flxInfo(0)
    .Redraw = False
    .GridLines = flexGridInset
    .SelectionMode = flexSelectionFree
    .Rows = 7   'iAnzInfoZeilen%
    .FixedRows = 0
    .Cols = 7
  
    For i% = 0 To 6
        .ColAlignment(i%) = flexAlignRightCenter
        If (i% Mod 2) Then
        Else
'            .ColAlignment(i%) = flexAlignLeftCenter
            .FillStyle = flexFillRepeat
            .col = i%
            .ColSel = i%
            .row = 0
            .RowSel = .Rows - 1
            .CellBackColor = vbWhite
            .FillStyle = flexFillSingle
        End If
    Next i%
    
    .TextMatrix(0, 1) = "Kunden m. Prämie"
    .TextMatrix(1, 1) = "Kunden mit Zusatzverkauf"
    .TextMatrix(2, 1) = "Kunden mit Sonderprämie"
    .TextMatrix(3, 1) = "Privatkunden"
    .TextMatrix(4, 1) = "Rezeptkunden"
    .TextMatrix(5, 1) = "Anzahl Rabatte"
    .TextMatrix(6, 1) = "Rabattsumme"
    
    .TextMatrix(0, 3) = "Summe Normtage"
    .TextMatrix(1, 3) = "Kunden/Normtag"
    .TextMatrix(2, 3) = "Rezepte/Normtag"
    .TextMatrix(3, 3) = "Sonderpräm.Kunden/Normtag"
    .TextMatrix(4, 3) = "Zusatzverk.Kunden/Normtag"
    .TextMatrix(5, 3) = "% Zusatzverk.Kunden/Rezeptkunden"
    .TextMatrix(6, 3) = "durchschn. Zusatzverkauf"
    
    .TextMatrix(0, 5) = "durchschn. Barverkauf/Privatkunde"
    .TextMatrix(1, 5) = "Prämienbasis/Kunde"
    .TextMatrix(2, 5) = "PrämienBasis/PrämienKunde"
    .TextMatrix(3, 5) = "Prämie"
    .TextMatrix(4, 5) = "Zusatzprämie"
    .TextMatrix(5, 5) = "Sonderprämie"
    .TextMatrix(6, 5) = "Prämiensumme"

    
'    For i% = 0 To 2
'        For j% = 1 To iAnzInfoZeilen%
'            h$ = iDatei$(j% - 1, i%)
'            If (h$ <> "") Then
'                txWert$ = clsDat.HoleDateiValue(h$, iKurz$(j% - 1, i%))
'
''                tx$ = clsDat.FirstLettersUcase(iBezeichnung$(j% - 1, i%))
'                tx$ = iBezeichnung$(j% - 1, i%)
'
'                .TextMatrix(j% - 1, 2 * i% + 1) = tx$
'                .TextMatrix(j% - 1, 2 * i% + 2) = txWert$
'            Else
'                .TextMatrix(j% - 1, 2 * i% + 1) = ""
'                .TextMatrix(j% - 1, 2 * i% + 2) = ""
'            End If
'        Next j%
'    Next i%
    .Redraw = True
End With

Call DefErrPop
End Sub

Public Sub cboMitarbeiter_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboMitarbeiter_Click")
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
Dim i%, ind%, j%
Dim h$
Dim Offset As Long

HatKunden = False

With cboMitarbeiter
  For i% = 1 To UBound(MitArb$)
    If Trim(Left(MitArb$(i%), 20)) = .List(.ListIndex) Then
      ind% = Val(Mid(MitArb$(i%), 22, 2))
      Exit For
    End If
  Next i%
End With

With flxarbeit(0)
  .Redraw = False
  .Rows = 0
  For i% = 0 To aTage% - 1
    h$ = aDetail$(0, ind%, i%) + vbTab
    If aDetail$(0, ind%, i%) = "" Then
      If i% = 0 Then
        h$ = ""
      Else
        'h$ = CStr(i%)
        h$ = aDetail$(0, ind%, i%)
      End If
      If Not Vorschau Then .AddItem h$
    Else
      h$ = ""
      For j% = 0 To UBound(aDetail$, 1)
        h$ = h$ + aDetail$(j%, ind%, i%) + vbTab
      Next j%
      .AddItem h$
    End If
  Next i%
  If .Rows > 0 Then
  .FixedRows = 1
  .FixedCols = 1
  End If
  .Redraw = True
End With
With picSummenzeile
  .Cls
  For i% = 1 To UBound(aDetail$, 1)
    Offset = 0
    For j% = 0 To i%
      Offset = Offset + flxarbeit(0).ColWidth(j%)
    Next j%
    .CurrentX = Offset - .TextWidth(aDetail$(i%, ind%, aTage%))
    .CurrentY = 0
    picSummenzeile.Print aDetail$(i%, ind%, aTage%)
  Next i%
  If xVal(aDetail$(4, ind%, aTage%)) > 0 Then HatKunden = True
End With

With flxInfo(0)
  For i% = 0 To .Rows - 1
    For j% = 0 To 5
      .TextMatrix(i%, j% + 1) = aInfo$(i%, j%, ind%)
    Next j%
  Next i%
End With

Call ZeilenEinAus

Call DefErrPop
End Sub


