VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAbholerVerwaltung 
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
   Icon            =   "AbholerVerwaltung.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   7890
   ScaleWidth      =   11295
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   5400
      Top             =   120
   End
   Begin VB.PictureBox picAnimationBack 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   1440
      ScaleHeight     =   2370
      ScaleWidth      =   5625
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   5655
      Begin ComCtl2.Animation aniAnimation 
         Height          =   1095
         Left            =   2280
         TabIndex        =   16
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
         Left            =   2040
         TabIndex        =   17
         Top             =   -240
         Width           =   5355
      End
   End
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
         TabIndex        =   13
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
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":09EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":0D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":1022
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":12B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":15CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":18E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":21AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":2440
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":2964
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":2BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":2E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":311A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":33AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":363E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":38D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":3B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":3E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":4196
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":4DE8
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
            Picture         =   "AbholerVerwaltung.frx":5A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":5B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":5DDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":5EF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":620A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":649C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":672E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":6A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":6D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":707C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":718E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":7420
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":76B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":7944
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":7A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":7B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":7DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":808C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":819E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":8430
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":8542
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":885C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":8B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "AbholerVerwaltung.frx":8EC8
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
Attribute VB_Name = "frmAbholerVerwaltung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "AbholerVerwaltung"
Const INFO_SECTION = "Infobereich AbholerVerwaltung"

Const ABHOLER_TEE = 0
Const ABHOLER_HOM = 1
Const ABHOLER_REZ = 2
Const ABHOLER_BESORGER = 3
Const ABHOLER_POST = 4
Const ABHOLER_ALLE = 5

Const NICHTS_TUN = 0
Const AKTUELL = 1
Const IN_ARBEIT = 2
Const FERTIG = 3
Const GELOESCHT = 4
Const BEREITS_WEG = 6


Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

'Dim InRowColChange%
Dim HochfahrenAktiv%
Dim ProgrammModus%
    
Dim ArtikelStatistik%

Dim AbholerAnzeige%

Private Const DefErrModul = "ABHOLERVERWALTUNG.FRM"

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

Private Sub cmdEsc_Click(index As Integer)
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

Unload Me

Call DefErrPop
End Sub

Private Sub cmdToolbar_Click(index As Integer)
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
If (index = 0) Then
'    Me.WindowState = vbMinimized
ElseIf (index <= 8) Then
    Call mnuBearbeitenInd_Click(index - 1)
ElseIf (index <= 16) Then
    Call mnuBearbeitenInd_Click(index)
ElseIf (index = 19) Then
    Call mnuBeenden_Click
End If

Call DefErrPop
End Sub

Private Sub flxarbeit_DragDrop(index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(flxarbeit(index), picBack(index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub flxArbeit_KeyPress(index As Integer, KeyAscii As Integer)
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

If (index = 0) Then
    If (KeyAscii = vbKeySpace) Then
        Call SetzePreisZeile(flxarbeit(0).row)
'        Call NaechsteBestellZeile
    ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
        Call SelectZeile(UCase(Chr$(KeyAscii)))
    End If
End If

Call DefErrPop
End Sub

Private Sub flxArbeit_DblClick(index As Integer)
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
cmdOk(index).Value = True
Call DefErrPop
End Sub

Private Sub flxarbeit_LostFocus(index As Integer)
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

Private Sub flxInfo_DblClick(index As Integer)
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
If (index = 0) Then
    cmdOk(0).Value = True
End If

Call DefErrPop
End Sub

Private Sub flxInfo_DragDrop(index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(flxInfo(index), picBack(index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub flxInfo_GotFocus(index As Integer)
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

If (index = 0) Then
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

Private Sub flxarbeit_GotFocus(index As Integer)
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

Private Sub flxarbeit_RowColChange(index As Integer)
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

If (index = 0) Then
'    If ((ProgrammChar$ = "B") And (flxarbeit(0).redraw = True) And (KeinRowColChange% = False)) Then
    If ((flxarbeit(0).Redraw = True) And (HochfahrenAktiv% = False)) Then
'        Call HighlightZeile
        Call FormKurzInfo
        flxInfo(0).row = 0
        flxInfo(0).col = 0
        
'        picQuittieren.Visible = False
'        If (flxarbeit(0).col = 2) Then
'            With picQuittieren
'                .Font.Name = wpara.FontName(0)
'                .Font.Size = wpara.FontSize(0)
'                .Width = flxarbeit(0).Width
'                .Height = flxarbeit(0).Height
'                .Cls
'                .CurrentY = 90
'                .CurrentX = 90
'                h$ = flxarbeit(0).TextMatrix(flxarbeit(0).row, 2)
'                picQuittieren.Print h$
'                .Width = TextWidth(h$) + 300
'                .Height = .CurrentY + 150
'                .Top = picBack(0).Top + flxarbeit(0).Top + flxarbeit(0).RowPos(flxarbeit(0).row) + flxarbeit(0).RowHeight(0)
'                .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(2)
'                .Visible = True
'            End With
'        End If
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

Caption = "Abholer-Verwaltung"

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

Call InitAnimation

With flxarbeit(0)
    .Cols = 17
    .Rows = 0
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|^ |>Abhol#|>Datum|>Uhrzeit|<Name|>Menge|^Meh|>BM|>Preis|^Stat|^Rez|^LS||||"
    .Rows = 1
End With
        
HochfahrenAktiv% = False
picBack(0).Visible = True

Call DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(lblArbeit(index), picBack(index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub lblInfo_DragDrop(index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(lblInfo(index), picBack(index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub mnuBearbeitenInd_Click(index As Integer)
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

Select Case index

    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call FormKurzInfo
            End If
        End If
    
    Case MENU_F3
            AbholerAnzeige% = (AbholerAnzeige% + 1) Mod 3
            Call AuslesenAbholer
    
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
    
    Case MENU_F7
        Call AuslesenAbholer
        
    Case MENU_SF3
        frmEditAbholer.Show 1
        
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

Private Sub mnuToolbarPositionInd_Click(index As Integer)
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

opToolbar.Position = index

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

Private Sub picBack_DragDrop(index As Integer, Source As Control, x As Single, Y As Single)
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
Call opToolbar.Move(picBack(index), picBack(index), Source, x, Y)
Call DefErrPop
End Sub

Private Sub picBack_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
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

Public Sub cmdOk_Click(index As Integer)
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
Dim row%, iRow%, iCol%, KommWeg%
Dim pzn$, h$
    
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


KommWeg% = True
h$ = AbholerKommentar$
If (h$ <> "") Then
    Call AnzeigeKommentar(h$, 6)
    KommWeg% = False
End If
If (KommWeg%) Then picQuittieren.Visible = False
    
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
    .Redraw = False
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
    
    .Redraw = True
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
        ch$ = Left$(.TextMatrix(i%, 5), 1)
        If (ch$ = SearchLetter$) Then
            gef% = i%
            Exit For
        End If
    Next i%
    
    If (gef% < 0) Then
        For i% = 1 To (.row - 1)
            ch$ = Left$(.TextMatrix(i%, 5), 1)
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
    .ColWidth(2) = TextWidth("999999")
    .ColWidth(3) = TextWidth("99.99.9999")
    .ColWidth(4) = TextWidth("99:9999")
    .ColWidth(5) = 0
    .ColWidth(6) = TextWidth("XXXXXX")
    .ColWidth(7) = TextWidth("XXX")
    .ColWidth(8) = TextWidth("9999")
    .ColWidth(9) = TextWidth("99999.99")
    .ColWidth(10) = TextWidth("XXXX")
    .ColWidth(11) = TextWidth("XXX")
    .ColWidth(12) = TextWidth("XX")
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = 0
    .ColWidth(16) = wpara.FrmScrollHeight
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
    .ColWidth(5) = .Width - spBreite%
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
Dim h$, h2$, Key$
    
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
           
Call ZeigeAbholerDaten

Call DefErrPop
End Sub
      
Sub AuslesenAbholer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenAbholer")
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
Dim i%, j%, erg%, KistenNr%, BesorgerStatus%, BesorgerMenge%, AnzBesorger%
Dim BesorgerPzn$, h$
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

IstAktiv% = True

h$ = ProgrammNamen$(ProgrammTyp%) + " - Sortierung nach "
If (AbholerAnzeige% = 0) Then
    h$ = h$ + "Gruppen/Datum"
ElseIf (AbholerAnzeige% = 1) Then
    h$ = h$ + "Abhol-Nummern"
Else
    h$ = h$ + "Artikel"
End If
Caption = h$

Call StartAnimation(Me, "Abholer werden eingelesen ...")

KeinRowColChange% = True

With flxarbeit(0)
    .Redraw = False
    .Rows = 1
    .row = 0
End With

AnzBesorger% = 0

'Call EinlesenWaMaBesorger

Kiste.OpenDatei

For KistenNr% = 1 To 999
    If (Kiste.Belegt(KistenNr%)) Then
        Kiste.GetKiste (KistenNr%)
        
        For i% = 0 To 9
            erg% = Kiste.GetInhalt(i%)
            If (erg%) Then
                If (Kiste.WasTun = "B") Then
                    BesorgerStatus% = Kiste.Status
                    BesorgerMenge% = 1
                    For j% = 1 To 10
                        h$ = RTrim$(Kiste.InfoText(j% - 1))
                        If Mid$(h$, 18, 3) = " x " Then
                            BesorgerMenge% = -Val(Mid$(h$, 21, 4))
                        End If
                        If Mid$(h$, 14, 4) = "PZN=" Then
                            BesorgerPzn$ = Mid$(h$, 18, 7)
                        End If
                    Next j%
                Else
                    BesorgerPzn$ = "9999999"
                    BesorgerMenge% = 1
                End If
                    
                With flxarbeit(0)
                    AnzBesorger% = AnzBesorger% + 1
            
                    .AddItem " "
'                    If (.row >= (.Rows - 1)) Then .Rows = .Rows + 1
                    .row = .Rows - 1
                    Call ZeigeAbholerZeile(KistenNr%, i%, BesorgerPzn$, BesorgerMenge%)
                End With
            End If
        Next i%
    End If
Next KistenNr%

Kiste.CloseDatei


With flxarbeit(0)
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = "XXXXXXX"
        .TextMatrix(1, 1) = " "
        For i% = 3 To (.Cols - 1)
            .TextMatrix(1, i%) = ""
        Next i%
        .TextMatrix(1, 5) = "Keine anzuzeigenden Abholer gespeichert !"
    Else
        .Rows = .row + 1
    End If
    .row = 1
    .col = 13
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = .col
    .Sort = 5
    .Redraw = True
    .TopRow = 1
    .row = 1
    .col = 5
    .ColSel = .col
    .SetFocus
End With

KeinRowColChange% = False
    
Call StopAnimation(Me)

'Call ZeigeWerte

IstAktiv% = False

Call DefErrPop
End Sub

Sub ZeigeAbholerZeile(KistenNr%, KistenInd%, pzn$, bm%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeAbholerZeile")
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
Dim i%, j%, lief%, l%, row%, col%, iBold%, iItalic%, FirstRed%, xAktiv%, ind%
Dim BackDunkel%
Dim lBack&
Dim EK#, Rabatt!, bmo!
Dim h$, h2$, s$, LiefName$, ArtName$, ArtMenge$, ArtMeh$, VonWo$, nm$, zusatz$, KontrollChar$, KontrollChar2$
Dim ActKontrollen$, actzuordnung$, SRT$, SQLStr$, ZusInfo$
Dim ArtikelKz As Byte
Dim BesorgerTxt$, WamaTxt$

BesorgerTxt$ = ""

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF) Then
    BesorgerTxt$ = "Rezeptur"
Else
    Call Taxe2ast(pzn$)
    BesorgerTxt$ = ast.kurz + ast.meng + ast.meh
End If


If (AbholerAnzeige% = 0) Then
    h$ = Format(Kiste.KistenStatus, "0")
    h$ = h$ + CVDatum(Left$(Kiste.VonWann, 2))
    h$ = h$ + Format(Asc(Mid$(Kiste.VonWann, 3, 1)), "00") + Format(Asc(Mid$(Kiste.VonWann, 4, 1)), "00")
    h$ = h$ + Format(KistenNr%, "0") + Format(KistenInd%, "0")
ElseIf (AbholerAnzeige% = 1) Then
    h$ = Format(KistenNr%, "000") + Format(KistenInd%, "0")
Else
    h$ = BesorgerTxt$
    h$ = h$ + Format(KistenNr%, "000") + Format(KistenInd%, "0")
End If
SRT$ = h$

ArtName$ = BesorgerTxt$
ArtMenge$ = ""
ArtMeh$ = ""
    
h$ = RTrim(Left$(BesorgerTxt$, 33))
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
ArtMeh$ = Mid$(BesorgerTxt$, 34)

KontrollChar$ = " "
If (Kiste.KistenStatus >= 3) Then
    KontrollChar$ = Chr$(214)
End If

KontrollChar2$ = " "
If (Kiste.Status >= 3) Then
    KontrollChar2$ = Chr$(214)
End If

'FirstRed% = False
'If (ww.fixiert = "2") Then
'    KontrollChar$ = Chr$(214)
'ElseIf (ActKontrollen$ <> "") Then
'    KontrollChar$ = "?"
'    If (ww.zukontrollieren = "1") Then
'        FirstRed% = True
'    End If
'End If

    
With flxarbeit(0)

    col% = .col

    iItalic% = False
    
    lBack& = vbButtonFace
    
    BackDunkel% = 0
    If (Kiste.KistenStatus = 4) Then
        BackDunkel% = True
    End If

    If (BackDunkel%) Then lBack& = wpara.FarbeDunklerBereich ' FarbeGray&        'vbGrayText

    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellFontItalic = iItalic%
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle

'    'damit Berechnung der Werte richtig
'    If (ww.zugeordnet = "N") Then
'        lBack& = wpara.FarbeDunklerBereich  ' FarbeGray& ' vbGrayText
'    End If

    .col = 1
    .CellFontName = "Symbol"
    If (FirstRed%) Then
        .CellBackColor = vbRed
    End If

    .col = 10
    .CellFontName = "Symbol"
    
    .col = 12
    .CellFontBold = True

    row% = .row
    .TextMatrix(row%, 0) = pzn$
    .TextMatrix(row%, 1) = KontrollChar$
    
    .TextMatrix(row%, 2) = Format(KistenNr%, "0")
    
    h$ = CVDatum(Left$(Kiste.VonWann, 2))
    .TextMatrix(row%, 3) = Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
    
    h$ = Kiste.VonWann
    .TextMatrix(row%, 4) = Format(Asc(Mid$(h$, 3, 1)), "00") + ":" + Format(Asc(Mid$(h$, 4, 1)), "00")
    
    .TextMatrix(row%, 5) = " " + ArtName$
    .TextMatrix(row%, 6) = ArtMenge$
    .TextMatrix(row%, 7) = ArtMeh$
    .TextMatrix(row%, 8) = Abs(bm%)
    .TextMatrix(row%, 9) = Format(Kiste.Preis, "0.00")
    
    .TextMatrix(row%, 10) = KontrollChar2$
    
    h$ = ""
    If (Kiste.RezeptNr > 0) Then
        h$ = "!"
    End If
    .TextMatrix(row%, 11) = h$
    
    h$ = ""
    If (Kiste.HatLieferSchein) Then
        h$ = "!"
    End If
    .TextMatrix(row%, 12) = h$
    
    .TextMatrix(row%, 13) = SRT$
    .TextMatrix(row%, 14) = Format(KistenNr%, "0")
    .TextMatrix(row%, 15) = Format(KistenInd%, "0")
    
    If (.TextMatrix(row%, 14) = .TextMatrix(row% - 1, 14)) Then
        .TextMatrix(row%, 2) = ""
        .TextMatrix(row%, 3) = ""
        .TextMatrix(row%, 4) = ""
    End If
    
    .col = col%
End With

Call DefErrPop
End Sub


Function KorrPzn$(Optional iPzn$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KorrPzn$")
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
Dim ch$

If (iPzn$ = "") Then
    With flxarbeit(0)
        iPzn$ = .TextMatrix(.row, 0)
    End With
End If

ch$ = Left$(iPzn$, 1)
If (Asc(ch$) > 127) Then
    ch$ = Chr$(Asc(ch$) - 128)
    Mid$(iPzn$, 1, 1) = ch$
End If

KorrPzn$ = iPzn$

Call DefErrPop
End Function

Function KorrTxt$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KorrTxt$")
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

With flxarbeit(0)
    h$ = Trim(.TextMatrix(.row, 5)) + "  " + Trim(.TextMatrix(.row, 6)) + .TextMatrix(.row, 7)
End With

KorrTxt$ = h$

Call DefErrPop
End Function

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

tmrStart.Enabled = False
Call AuslesenAbholer

Call DefErrPop
End Sub

Sub ZeigeAbholerDaten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeAbholerDaten")
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
Dim j%, erg%, row%, KistenNr%, KistenInd%, BesorgerStatus%, BesorgerMenge%
Dim h$, BesorgerPzn$

With flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 14))
    KistenNr% = Val(h$)
    h$ = RTrim$(.TextMatrix(row%, 15))
    KistenInd% = Val(h$)
    
    KeinRowColChange% = True
    
    If (Trim(.TextMatrix(row%, 5)) = "Rezeptur") Then
        MagSpeicherIndex% = 0
        AnfMagIndex& = KistenNr% * 100& + KistenInd% ' Kiste.ErstMagSatz
        
        frmTaxieren.Show 1
    Else
        Call BesorgerInfo(KistenNr%, KorrPzn$, KorrTxt$)
    End If
    
    Kiste.OpenDatei
    If (Kiste.Belegt(KistenNr%)) Then
        Kiste.GetKiste (KistenNr%)
        
        erg% = Kiste.GetInhalt(KistenInd%)
        If (erg%) Then
            If (Trim(.TextMatrix(row%, 5)) = "Rezeptur") Then
                BesorgerPzn$ = "9999999"
                BesorgerMenge% = 1
                Kiste.Preis = AnfMagPreis#
                erg% = Kiste.PutInhalt(KistenInd%)
            ElseIf (Kiste.WasTun = "B") Then
                BesorgerStatus% = Kiste.Status
                BesorgerMenge% = 1
                For j% = 1 To 10
                    h$ = RTrim$(Kiste.InfoText(j% - 1))
                    If Mid$(h$, 18, 3) = " x " Then
                        BesorgerMenge% = -Val(Mid$(h$, 21, 4))
                    End If
                    If Mid$(h$, 14, 4) = "PZN=" Then
                        BesorgerPzn$ = Mid$(h$, 18, 7)
                    End If
                Next j%
                
            End If
            
            Call ZeigeAbholerZeile(KistenNr%, KistenInd%, BesorgerPzn$, BesorgerMenge%)
        End If
    End If
    Kiste.CloseDatei
    
    .col = 5
    .ColSel = .col
    
    KeinRowColChange% = False
End With

Call DefErrPop
End Sub

Function AbholerKommentar$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerKommentar$")
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
Dim j%, erg%, row%, KistenNr%, KistenInd%, AbInd%
Dim h$, ret$

ret$ = ""

With flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 14))
    KistenNr% = Val(h$)
    h$ = RTrim$(.TextMatrix(row%, 15))
    KistenInd% = Val(h$)

    If (Trim(.TextMatrix(row%, 5)) = "Rezeptur") Then
        AbInd% = 0
    Else
        AbInd% = 3
    End If
End With
    

Kiste.OpenDatei
If (Kiste.Belegt(KistenNr%)) Then
    Kiste.GetKiste (KistenNr%)
    
    erg% = Kiste.GetInhalt(KistenInd%)
    If (erg%) Then
        For j% = AbInd% To 9
            h$ = Trim(Kiste.InfoText(j%))
            If (AbInd% = 0) And (Left$(h$, 3) = "240") Then
                h$ = ""
            End If
            If (h$ <> "") Then
                ret$ = ret$ + h$ + Chr$(13)
            End If
        Next j%
    End If
End If
Kiste.CloseDatei

AbholerKommentar$ = ret$

Call DefErrPop
End Function

Sub AnzeigeKommentar(h$, pos%, Optional modus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AnzeigeKommentar")
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
Dim i%, wi%, MaxWi%, ind%, sperren%, spBreite%(3), AnzR%
Dim lNr&
Dim h2$

With picQuittieren
    .Visible = False
    .Font.Name = wpara.FontName(0)
    .Font.Size = wpara.FontSize(0)
    .Width = flxarbeit(0).Width
    .Height = flxarbeit(0).Height
    .Cls
    .CurrentY = 90
    MaxWi% = 0
    If (modus% = 1) Then
        spBreite%(0) = 90
        spBreite%(1) = spBreite%(0) + TextWidth("99.99.9999  ")
        spBreite%(2) = spBreite%(1) + TextWidth("WWWWWW(999)  ")
        spBreite%(3) = spBreite%(2) + TextWidth("9999  ")
    End If
    AnzR% = 0
    Do
        If (h$ = "") Or (AnzR% >= 10) Then
            Exit Do
        End If
        AnzR% = AnzR% + 1
        ind% = InStr(h$, vbCr)
        If (ind% > 0) Then
            h2$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Else
            h2$ = h$
            h$ = ""
        End If
        
        .CurrentX = 90
        If (modus% = 1) Then
            For i% = 0 To 5
                ind% = InStr(h2$, "@")
                If (ind% > 0) Then
                    .CurrentX = spBreite%(i%)
                    picQuittieren.Print Left$(h2$, ind% - 1);
                    h2$ = Mid$(h2$, ind% + 1)
                    If (i% = 2) Then wi% = spBreite%(3) + TextWidth(h2$)
                Else
                    picQuittieren.Print
                    Exit For
                End If
            Next i%
            If (wi% > MaxWi%) Then
                MaxWi% = wi%
            End If
        Else
            wi% = TextWidth(h2$)
            If (wi% > MaxWi%) Then
                MaxWi% = wi%
            End If
            picQuittieren.Print h2$
        End If
    Loop
    .Width = MaxWi% + 300
    .Height = .CurrentY + 150
    If (modus% = 2) Then
        .Top = picBack(0).Top + flxarbeit(0).Top + flxarbeit(0).RowPos(flxarbeit(0).row) + flxarbeit(0).RowHeight(0)
        .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(pos%)
    Else
        .Top = picBack(0).Top + flxarbeit(0).Top + flxarbeit(0).RowPos(flxarbeit(0).row)
        .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(pos%) - .Width
    End If
    .Visible = True
End With

'sperren% = True
'If (modus% = 2) Then sperren% = False
sperren% = False
'If (modus% = 0) And (vAnzeigeSperren%) Then sperren% = True

'If (sperren%) Then
'    ReDim Preserve KommentarOk&(UBound(KommentarOk&) + 1)
'    KommentarOk&(UBound(KommentarOk)) = lNr&
'    Call SetKommentarTyp(0)
'End If


Call DefErrPop
End Sub


