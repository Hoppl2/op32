VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmArtikelZuordnung 
   Caption         =   "Matchcode - Auswahl"
   ClientHeight    =   7350
   ClientLeft      =   285
   ClientTop       =   600
   ClientWidth     =   13245
   ScaleHeight     =   7350
   ScaleWidth      =   13245
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   12120
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picToolbarOrg 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
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
      Left            =   11880
      ScaleHeight     =   360
      ScaleWidth      =   1095
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   10440
      Top             =   3720
   End
   Begin VB.ListBox lstSortierung 
      Height          =   255
      Left            =   10320
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
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
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picSave 
      Height          =   1095
      Left            =   10440
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   8
      Top             =   480
      Width           =   975
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
      Left            =   6600
      ScaleHeight     =   300
      ScaleWidth      =   2355
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8160
      Index           =   0
      Left            =   0
      ScaleHeight     =   8160
      ScaleWidth      =   10095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   10095
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   10  'Stift maskieren
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   960
         ScaleHeight     =   555
         ScaleWidth      =   1275
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   450
         Index           =   0
         Left            =   3360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1200
      End
      Begin VB.CommandButton cmdEsc 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   450
         Index           =   0
         Left            =   5040
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfoZusatz 
         Height          =   780
         Index           =   0
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4440
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
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   1500
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   960
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
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   3960
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
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
         SelectionMode   =   1
      End
      Begin ComctlLib.ImageList imgToolbar 
         Index           =   4
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   64
         ImageHeight     =   64
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   16
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":0852
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":10A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":40F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":4948
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":519A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":81EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":8A3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":BA90
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":C2E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":CB34
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":D386
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":DBD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":10C2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":1147C
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ArtikelZuordnung.frx":11CCE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblArbeit 
         Alignment       =   2  'Zentriert
         Caption         =   "Manuelles Erfassen"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
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
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   9615
      End
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   2
      Left            =   720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":14D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":16172
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":175C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":18A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":19E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":1B2BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":1E30C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":1F75E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":227B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":25802
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":28854
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":2B8A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":2E8F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3194A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":32D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":341EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   3
      Left            =   120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":37240
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":37A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":382E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3B336
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3BB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3C3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3F42C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":3FC7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":42CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":43522
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":43D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":445C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":44E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":47E6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":486BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":48F0E
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
            Picture         =   "ArtikelZuordnung.frx":4BF60
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4C072
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4C304
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4C416
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4C528
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4C7BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4CA4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4CB5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4CDF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4CF02
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4D014
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4D2A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4D538
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4D7CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4DA5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4DCEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4DE00
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4DF12
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4E024
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4E2B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4E3C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4E6E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4E9FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4ED16
            Key             =   ""
         EndProperty
      EndProperty
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
            Picture         =   "ArtikelZuordnung.frx":4F030
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4F2C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4F5DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4F8F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4FA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4FC9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":4FF2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":5003E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":50150
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":503E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":50674
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":50906
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":50B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":50E2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":510BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":5134E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":51460
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":516F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":51984
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":51C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":51EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":521C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":524DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ArtikelZuordnung.frx":527F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDateiInd 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuDummy10 
         Caption         =   "-"
      End
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
         Caption         =   ""
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
         Caption         =   ""
         Index           =   15
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
         Index           =   16
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuDummy 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuBearbeitenLayout 
         Caption         =   "La&yout editieren"
      End
      Begin VB.Menu mnuDummy1 
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
      Begin VB.Menu mnuDummy2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuZusatzInfo 
         Caption         =   "Artikel-S&tatistik"
      End
   End
End
Attribute VB_Name = "frmArtikelZuordnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "ArtikelZuordnung"
Const INFO_SECTION = "Infobereich ArtikelZuordnung"
Const ARBEIT_SECTION = "Arbeitsbereich ArtikelZuordnung"


Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich
'Dim ArbeitMain As clsArbeitBereich

Dim HochfahrenAktiv%

Dim ArtikelStatistik%

Dim ProgrammModus%

Dim FabsErrf%
Dim FabsRecno&

Dim INI_DATEI As String

Dim ZuordLief%, ZuordBevorratungsZeit%, ZuordAbBm%
Dim SortModus%

Dim ZuordSortStr$(2)

Private Const DefErrModul = "ARTIKELZUORDNUNG.FRM"

Public Sub WechselModus(NeuerModus%, Optional NeuMachen% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WechselModus")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ind%, erg%
Dim h$

KeinRowColChange% = True
Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        
        mnuBearbeitenInd(MENU_F2).Enabled = True
        mnuBearbeitenInd(MENU_F3).Enabled = True
        mnuBearbeitenInd(MENU_F4).Enabled = True
        mnuBearbeitenInd(MENU_F5).Enabled = True
        mnuBearbeitenInd(MENU_F6).Enabled = True 'false
        mnuBearbeitenInd(MENU_F7).Enabled = False
        mnuBearbeitenInd(MENU_F8).Enabled = True
        mnuBearbeitenInd(MENU_F9).Enabled = True
        mnuBearbeitenInd(MENU_SF2).Enabled = False
        mnuBearbeitenInd(MENU_SF3).Enabled = True
        mnuBearbeitenInd(MENU_SF4).Enabled = True
        mnuBearbeitenInd(MENU_SF5).Enabled = False
        mnuBearbeitenInd(MENU_SF6).Enabled = False
        mnuBearbeitenInd(MENU_SF7).Enabled = False
        mnuBearbeitenInd(MENU_SF8).Enabled = False
        
        cmdOk(0).default = True
        cmdEsc(0).Cancel = True
        
        flxarbeit(0).SelectionMode = flexSelectionByRow
        flxarbeit(0).col = 0
        flxarbeit(0).ColSel = flxarbeit(0).Cols - 1
        
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
        mnuBearbeitenInd(MENU_F3).Enabled = True
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
        
        cmdOk(0).default = True
        cmdEsc(0).Cancel = True

'        flxarbeit(0).SelectionMode = flexSelectionFree
'        flxarbeit(0).col = flxarbeit(0).Cols - 1
'        flxarbeit(0).ColSel = flxarbeit(0).col
        
        flxarbeit(0).BackColorSel = vbMagenta
        
        With flxInfo(0)
            .BackColorSel = vbMagenta
            .row = 0
            .col = 1
            .SetFocus
        End With
        
        h$ = Me.Caption
        Me.Caption = h$ + " (EDITIER-MODUS)"

End Select
KeinRowColChange% = False

For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%
    
ProgrammModus% = NeuerModus%

Call clsError.DefErrPop
End Sub

Private Sub cmdEsc_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdEsc_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, h$

h$ = ""
'Match1.MatchcodeErg = h$
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdOk_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdOk_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%, erg%, row%, col%
Dim lInd&
Dim h$, h2$, pzn$, txt$

If (ProgrammModus% = 1) Then
    If (ActiveControl.Name = flxInfo(0).Name) Then
        With flxInfo(0)
            row% = .row
            col% = .col
            h$ = RTrim(.text)
        End With
        If (col% Mod 2) Then
            MatchTyp% = -1
            Call InfoMain.EditInfoBelegung
            Call AuswahlKurzInfo
        End If
    End If
ElseIf (ActiveControl.Name = flxarbeit(0).Name) Then
    If (ActiveControl.index = 1) Then
        Call Match1.EditSatz
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdToolbar_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdToolbar_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index <= 8) Then
    Call mnuBearbeitenInd_Click(index - 1)
ElseIf (index <= 16) Then
    Call mnuBearbeitenInd_Click(index)
ElseIf (index = 19) Then
'    Call mnuBeenden_Click
End If

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_DblClick")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
cmdOk(0).Value = True
Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_GotFocus")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call AuswahlKurzInfo

With flxarbeit(index)
    .HighLight = flexHighlightAlways
    .col = 0
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_KeyPress")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (KeyAscii = vbKeySpace) Then
    Call ArtikelZuordnung(flxarbeit(0).row)
End If

Call clsError.DefErrPop
End Sub

Private Sub ArtikelZuordnung(row%, Optional ZuordModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ArtikelZuordnung")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iLief%, arow%
Dim ch$, pzn$

With flxarbeit(0)
    pzn$ = .TextMatrix(row%, 0)

    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Ass1.GetRecord (FabsRecno& + 1)
        
        If ((ZuordModus% = 0) And (.TextMatrix(row%, 1) = "")) Or (ZuordModus% = 1) Then
            iLief% = ZuordLief%
            ch$ = Chr$(214)
        Else
            iLief% = 0
            ch$ = ""
        End If
        Ass1.lief = iLief%
        
        Ass1.PutRecord (FabsRecno& + 1)
        .TextMatrix(row%, 1) = ch$
    
        arow% = .row
        .FillStyle = flexFillRepeat
        KeinRowColChange% = True
        .row = row%
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        If (.TextMatrix(row%, 1) <> "") Then
            .CellForeColor = vbBlue
        Else
            .CellForeColor = .ForeColor
        End If
        
        KeinRowColChange% = False
        .row = arow%
        If (.row < (.Rows - 1)) Then .row = .row + 1
        .col = 0
        .ColSel = .Cols - 1
        
        .FillStyle = flexFillSingle
    End If
End With

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_LostFocus")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxarbeit(index)
    .HighLight = flexHighlightNever
End With

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_RowColChange(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_RowColChange")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

'If ((flxarbeit(0).Visible = True) And (KeinRowColChange% = False)) Then
If (KeinRowColChange% = False) And (HochfahrenAktiv% = False) Then
    Call AuswahlKurzInfo
    flxInfo(0).row = 0
    flxInfo(0).col = 0
End If

Call clsError.DefErrPop
End Sub

Private Sub flxInfo_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfo_DblClick")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
cmdOk(0).Value = True
Call clsError.DefErrPop
End Sub

Private Sub flxInfo_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfo_GotFocus")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ArbeitRow%, InfoRow%, arow%, aCol%, erg%
Dim stSatz&, st2Satz&
Dim pzn$

If (index = 0) Then
    ArbeitRow% = flxarbeit(0).row
    
    With flxInfo(0)
        arow% = .row
        aCol% = .col
        
        InfoRow% = 0
        
        For i% = InfoRow% To (.Rows - 1)
            .TextMatrix(i%, 0) = ""
        Next i%
        For i% = 0 To (.Rows - 1)
            .row = i%
            .col = 0
            .CellFontBold = True
            .CellForeColor = .ForeColor
        Next i%
        .row = arow%
        .col = aCol%
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyDown")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%
Dim h$

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
        picToolTip.Left = cmdToolbar(ind%).Left
        picToolTip.Top = 660
        picToolTip.Visible = True
        picToolTip.Cls
        picToolTip.CurrentX = 2 * Screen.TwipsPerPixelX
        picToolTip.CurrentY = 0
        picToolTip.Print h$
        KeyCode = 0
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Load")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, erg%
Dim l&
Dim h$

HochfahrenAktiv% = True

If (iNewLine) Then
    WindowState = ProjektForm.WindowState
    Width = ProjektForm.Width
    Height = ProjektForm.Height
    Top = ProjektForm.Top
    Left = ProjektForm.Left
    Call wPara1.ControlBorderless(Me, 3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
Else
    Width = Screen.Width - (1200 * wPara1.BildFaktor)
    Height = Screen.Height - (1200 * wPara1.BildFaktor)
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End If

'If (iNewLine) Then
'    WindowState = ProjektForm.WindowState
'    Width = ProjektForm.Width
'    Height = ProjektForm.Height
'    Top = ProjektForm.Top
'    Left = ProjektForm.Left
'Else
'    Width = Screen.Width - (800 * wPara1.BildFaktor)
'    Height = Screen.Height - (1200 * wPara1.BildFaktor)
'    Top = (Screen.Height - Height) / 2
'    Left = (Screen.Width - Width) / 2
'End If

INI_DATEI = CurDir + "\winop.ini"
   
Caption = "Artikel-Zuordnungen für " + Trim(frmLiefStammdaten.txtStammdaten(4).text)

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
cmdToolbar(8).ToolTipText = "F9"    ' Abmelden"
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

For i = MENU_SF2 To MENU_SF9
    mnuBearbeitenInd(i).Enabled = False
    cmdToolbar(i).Enabled = False
Next i


Call wPara1.InitFont(Me)
Call HoleIniWerte

'Set ArbeitMain = New clsArbeitBereich
'Call ArbeitMain.InitArbeitBereich(flxarbeit(0), INI_DATEI, ARBEIT_SECTION)

Set InfoMain = New clsInfoBereich
Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, INFO_SECTION, 2)
Call InfoMain.ZeigeInfoBereich("", False)
Call ZeigeInfoBereichAdd(0)
flxInfo(0).row = 0
flxInfo(0).col = 0

Set opBereich = New clsOpBereiche
Call opBereich.InitBereich(Me, opToolbar)
opBereich.AutoRedraw = 0
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = False
opBereich.ArbeitWasDarunter = False
opBereich.InfoTitel = False
opBereich.InfoZusatz = ArtikelStatistik%
opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = -2

mnuZusatzInfo.Checked = ArtikelStatistik%

ProgrammModus% = 0

'flxarbeit(0).row = 1
'Call WechselModus(0)

With flxarbeit(0)
'    .Cols = 24
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|||<Name|>Menge|^Meh|^Herst|>BM|>BMopt|>POS|>AEP|>AVP|^Rez|"
    .Rows = 1
End With
        
mnuBearbeitenZusatz(0).Caption = "&Alle Artikel zuordnen"
mnuBearbeitenZusatz(0).Enabled = True

Load mnuBearbeitenZusatz(1)
mnuBearbeitenZusatz(1).Caption = "Alle Artikel ab &BM " + Trim(frmLiefStammdaten.txtStammdaten3(0).text) + " zuordnen"
mnuBearbeitenZusatz(1).Enabled = True

Load mnuBearbeitenZusatz(2)
mnuBearbeitenZusatz(2).Caption = "Alle RX-Artikel ab optimaler BM ... zuordnen"
mnuBearbeitenZusatz(2).Enabled = True

Load mnuBearbeitenZusatz(3)
mnuBearbeitenZusatz(3).Caption = "Alle NonRX-Artikel ab optimaler BM ... zuordnen"
mnuBearbeitenZusatz(3).Enabled = True

Load mnuBearbeitenZusatz(4)
mnuBearbeitenZusatz(4).Caption = "Alle Zuordnungen &rücksetzen"
mnuBearbeitenZusatz(4).Enabled = True

HochfahrenAktiv% = False
picBack(0).Visible = True

ZuordSortStr$(0) = "Alphabet."
ZuordSortStr$(1) = "Herst + Alphabet."
ZuordSortStr$(2) = "Zuord + Alphabet"

Call clsError.DefErrPop
End Sub

Sub RefreshBereichsFlexSpalten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsFlexSpalten")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, spBreite%
Dim sp&
            
With flxarbeit(0)
    Font.Bold = True
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth("X")
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .ColWidth(4) = TextWidth("XXXXXX")
    .ColWidth(5) = TextWidth("XXX")
    .ColWidth(6) = TextWidth("XXXXXXX")
    .ColWidth(7) = TextWidth("99999.99 ")
    .ColWidth(8) = TextWidth("99999.99 ")
    .ColWidth(9) = TextWidth("99999.99 ")
    .ColWidth(10) = TextWidth("99999.99 ")
    .ColWidth(11) = TextWidth("99999.99 ")
    .ColWidth(12) = TextWidth("XXX")
    .ColWidth(13) = wPara1.FrmScrollHeight   '+ 2 * wPara1.FrmBorderHeight
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
    .ColWidth(3) = .Width - spBreite%
End With
    
With picProgress
    .Left = flxarbeit(0).Width / 4
    .Top = flxarbeit(0).Height / 2
    .Width = flxarbeit(0).Width / 2
    .Height = .TextHeight("99 %") + 180
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
    sp& = .Width / 15 + 15
    For i% = 0 To 14
        .ColWidth(i%) = sp&
    Next i%
End With
        
Call clsError.DefErrPop
End Sub

Sub RefreshBereichsControlsAdd()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsControlsAdd")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call clsError.DefErrPop
End Sub

Sub RefreshBereichsFarbenAdd()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsFarbenAdd")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call clsError.DefErrPop
End Sub

Public Sub AuswahlKurzInfo()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AuswahlKurzInfo")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim row%, iRow%, iCol%
Dim pzn$, ch$, ActKontrollen$, ActZusatz$, actzuordnung$

If (flxarbeit(0).Rows < 2) Then Call clsError.DefErrPop: Exit Sub
row% = flxarbeit(0).row
pzn$ = RTrim$(flxarbeit(0).TextMatrix(row%, 0))
If (pzn$ = "XXXXXXX") Then Call clsError.DefErrPop: Exit Sub
If (Len(pzn$) <> 7) Then Call clsError.DefErrPop: Exit Sub

row% = flxarbeit(0).row
iRow% = flxInfo(0).row
iCol% = flxInfo(0).col
Call InfoMain.ZeigeInfoBereich(pzn$, True)
If (ActiveControl.Name <> flxInfo(0).Name) Then
    Call ZeigeInfoBereichAdd(0)
End If
flxInfo(0).row = iRow%
flxInfo(0).col = iCol%

If (opBereich.InfoZusatz) Then
    Call ZeigeInfoZusatz(pzn$)
End If

'FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
'If (FabsErrf% = 0) Then
'    mnuBearbeitenInd(MENU_SF5).Enabled = True
'Else
'    mnuBearbeitenInd(MENU_SF5).Enabled = False
'End If
'cmdToolbar(12).Enabled = mnuBearbeitenInd(MENU_SF5).Enabled

Call clsError.DefErrPop
End Sub

Sub ZeigeInfoBereichAdd(index%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeInfoBereichAdd")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim j%

With flxInfo(index%)
    .Redraw = False
    
    .row = 0
    .col = 0
    .CellFontBold = True
    
    .TextMatrix(0, 0) = ZuordSortStr$(SortModus%)
    
    j% = 2
    Do While (j% <= InfoMain.AnzInfoZeilen%)
        .TextMatrix(j% - 1, 0) = ""
        j% = j% + 1
    Loop
    
    .Redraw = True
End With

Call clsError.DefErrPop
End Sub

Sub ZeigeInfoZusatz(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeInfoZusatz")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, Monat%, erg%
Dim Jahr&, AltJahr&, Termin&
Dim iWert!

With flxInfoZusatz(0)
    .Redraw = False
    
    If (iNewLine) Then
        .GridLines = flexGridFlat
        .BorderStyle = flexBorderNone
        
        .SelectionMode = flexSelectionFree
        .Rows = 2
        .Cols = 15
        .FixedRows = 0
        .FixedCols = 0
        
        For i% = 0 To 1
            .FillStyle = flexFillRepeat
            .row = i
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            If (i% = 0) Then
                .CellBackColor = RGB(199, 176, 123)
            Else
                .CellBackColor = RGB(232, 217, 172)
                .CellFontSize = .Font.Size + 2
                .CellFontBold = False
            End If
            .CellAlignment = flexAlignCenterCenter
            .FillStyle = flexFillSingle
        Next i%
    Else
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
    End If
    
    Set ArtStat1 = New clsArtStatistik
    erg% = ArtStat1.StatistikRechnen(pzn$)
    If (erg%) Then
        AltJahr& = -1
        j% = 0
        For i% = 0 To 12
            Termin& = ArtStat1.Anfang - i% - 1
            Jahr& = (Termin& - 1) \ 12
            Monat% = Termin& - Jahr& * 12
            If (AltJahr& <> Jahr&) Then
                .TextMatrix(0, j%) = Str$(Jahr&)
                iWert! = ArtStat1.JahresWert(Jahr&)
                If (iWert! <> 0!) Then
                    .TextMatrix(1, j%) = Str$(iWert!)
                Else
                    .TextMatrix(1, j%) = ""
                End If
                .row = 0
                .col = j%
                .CellFontBold = True
                .row = 1
                .CellFontBold = True
                j% = j% + 1
                AltJahr& = Jahr&
            End If
            
            .TextMatrix(0, j%) = Para1.MonatKurz(Monat%)
            
            iWert! = ArtStat1.MonatsWert(i% + 1)
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
    Set ArtStat1 = Nothing

    .Redraw = True
End With

Call clsError.DefErrPop

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_QueryUnload")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.SpeicherIniToolbar
Set opToolbar = Nothing
Set InfoMain = Nothing
Call clsError.DefErrPop
End Sub

Private Sub mnuBearbeitenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuBearbeitenInd_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, erg%, row%, col%, ind%, iAusr%, bm%
Dim bmo!
Dim l&
Dim pzn$, txt$, mErg$, h$

Select Case index
    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call AuswahlKurzInfo
            End If
        Else
            mErg$ = clsDialog.MatchCode(0, pzn$, txt$, False, False)
            If (mErg$ <> "") Then
                Do
                    If (mErg$ = "") Then Exit Do
                    
                    ind% = InStr(mErg$, vbTab)
                    h$ = Left$(mErg$, ind% - 1)
                    mErg$ = Mid$(mErg$, ind% + 1)
                    
                    If (Right$(h$, 1) = "-") Then
                        h$ = Left$(h$, Len(h$) - 1)
                    End If
                    ind% = InStr(h$, "@")
                    pzn$ = Left$(h$, ind% - 1)
                
                    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        Call Ass1.GetRecord(FabsRecno& + 1)
                        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            Ast1.GetRecord (FabsRecno& + 1)
                            
                            If (Ass1.opt <= 0) Then
                                bmo! = 0
                            Else
                                bmo! = Ass1.opt
'                                bmo% = Int(Ass1.opt + 0.501)
                            End If
                            bm% = clsOpTool.CalcDirektBM%(ZuordBevorratungsZeit%, bmo!, Ass1.poslag)
                
                            h$ = pzn$ + vbTab
                            If (Ass1.lief = ZuordLief%) Then h$ = h$ + Chr$(214)
                            h$ = h$ + vbTab
                            h$ = h$ + vbTab + Ast1.kurz + vbTab + Ast1.meng + vbTab + Ast1.meh
                            h$ = h$ + vbTab + Ast1.herst
                            h$ = h$ + vbTab + Format(bm%, "0")
                            h$ = h$ + vbTab + Format(Ass1.opt, "0.0")
                            h$ = h$ + vbTab + Format(Ass1.poslag, "0")
                            h$ = h$ + vbTab + Format(Ast1.aep, "0.00")
                            h$ = h$ + vbTab + Format(Ast1.AVP, "0.00")
                            h$ = h$ + vbTab + Ast1.rez
                            flxarbeit(0).AddItem h$
                            
                            If (flxarbeit(0).TextMatrix(1, 0) = "XXXXXXX") Then
                                flxarbeit(0).RemoveItem 1
                            End If
                        End If
                    End If
                Loop
                
                With flxarbeit(0)
                    .Redraw = False
                    .FillStyle = flexFillRepeat
                    .row = 1
                    .col = 1
                    .RowSel = .Rows - 1
                    .ColSel = .col
                    .CellFontName = "Symbol"
                
                    KeinRowColChange% = True
                    For i% = 1 To (.Rows - 1)
                        If (.TextMatrix(i%, 1) <> "") Then
                            .row = i%
                            .col = 0
                            .RowSel = .row
                            .ColSel = .Cols - 1
                            .CellForeColor = vbBlue
                        End If
                    Next i%
                    KeinRowColChange% = False
                    .FillStyle = flexFillSingle
                End With
                
                Call SortiereArtikelZuordnung
            End If
        End If
    
    Case MENU_F3
        SortModus% = (SortModus% + 1) Mod 3
        Call SortiereArtikelZuordnung

    Case MENU_F5
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call AuswahlKurzInfo
            End If
        End If
        
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
    
    Case MENU_SF9
'        If (ProgrammModus% = 0) Then
'            Call WechselModus(1)
'        Else
'            Call WechselModus(0)
'        End If

End Select

Call clsError.DefErrPop
End Sub

Private Sub mnuBearbeitenLayout_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuBearbeitenLayout_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:

If (ProgrammModus% = 0) Then
    Call WechselModus(1)
Else
    Call WechselModus(0)
End If

Call clsError.DefErrPop
End Sub

Private Sub mnuBearbeitenZusatz_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuBearbeitenZusatz_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ok%, zModus%, arow%
Dim ZuordAbBMopt#
Dim h$

If (index = 4) Then
    If (clsDialog.MessageBox("Alle Zuordnungen für diesen Lieferanten rücksetzen ?", vbYesNo Or vbDefaultButton2) <> vbYes) Then
        Call clsError.DefErrPop: Exit Sub
    End If
ElseIf (index >= 2) Then
    h$ = ""
    Do
        If (index = 2) Then
            h$ = Trim(clsDialog.MyInputBox("Alle RX-Artikel ab optimaler BM:", "Artikel zuordnen", h$))
        Else
            h$ = Trim(clsDialog.MyInputBox("Alle NonRX-Artikel ab optimaler BM:", "Artikel zuordnen", h$))
        End If
        If (h$ = "") Then
            Call clsError.DefErrPop: Exit Sub
        End If
        
        ZuordAbBMopt = clsOpTool.xVal(h$)
        If (ZuordAbBMopt >= 0) Then
            Exit Do
        End If
    Loop
End If


With flxarbeit(0)
    .Redraw = False
    arow% = .row
    For i% = 1 To (.Rows - 1)
        ok% = True
        If (index < 4) Then
            zModus% = 1
        Else
            zModus% = 2
        End If
        If (index = 1) Then
            If (Val(.TextMatrix(i%, 7)) < ZuordAbBm%) Then ok% = 0
        ElseIf (index = 2) Or (index = 3) Then
            ok = 0
            If (clsOpTool.xVal(.TextMatrix(i%, 8)) >= ZuordAbBMopt) Then
                h$ = Trim(.TextMatrix(i%, 12))
                If (index = 2) And (h$ = "+") Then
                    ok = True
                ElseIf (index = 3) And (h$ <> "+") Then
                    ok = True
                End If
            End If
        End If
        If (ok%) Then Call ArtikelZuordnung(i%, zModus%)
    Next i%
    .row = arow%
    .col = 0
    .ColSel = .Cols - 1
    .Redraw = True
End With

Call clsError.DefErrPop
End Sub

Private Sub mnuBeenden_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuBeenden_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Unload Me
Call clsError.DefErrPop
End Sub

Private Sub mnuToolbarGross_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuToolbarGross_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%

If (opToolbar.BigSymbols) Then
    opToolbar.BigSymbols = False
Else
    opToolbar.BigSymbols = True
End If

Call clsError.DefErrPop
End Sub

Private Sub mnuToolbarLabels_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuToolbarLabels_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (opToolbar.Labels) Then
    opToolbar.Labels = False
Else
    opToolbar.Labels = True
End If

Call clsError.DefErrPop
End Sub

Private Sub mnuToolbarPositionInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuToolbarPositionInd_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%

opToolbar.Position = index

Call clsError.DefErrPop
End Sub

Private Sub mnuToolbarVisible_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuToolbarVisible_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (opToolbar.Visible) Then
    opToolbar.Visible = False
    mnuToolbarVisible.Caption = "Einblenden"
Else
    opToolbar.Visible = True
    mnuToolbarVisible.Caption = "Ausblenden"
End If

Call clsError.DefErrPop
End Sub

Private Sub opToolbar_Resized()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("opToolbar_Resized")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opBereich.ResizeWindow
Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Resize")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error Resume Next

If (HochfahrenAktiv%) Then Call clsError.DefErrPop: Exit Sub

If (Me.WindowState = vbMinimized) Then Call clsError.DefErrPop: Exit Sub

Call opBereich.ResizeWindow

picSave.Visible = False

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_DragDrop")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.Move(flxarbeit(index), picBack(index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub flxInfo_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfo_DragDrop")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.Move(flxInfo(index), picBack(index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lblarbeit_DragDrop")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.Move(lblArbeit(index), picBack(index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub lblInfo_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lblInfo_DragDrop")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.Move(lblInfo(index), picBack(index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub picBack_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picBack_DragDrop")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call opToolbar.Move(picBack(index), picBack(index), Source, x, y)
Call clsError.DefErrPop
End Sub

'Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("picToolbar_MouseDown")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'picToolbar.Drag (vbBeginDrag)
'opToolbar.DragX = x
'opToolbar.DragY = Y
'Call clsError.DefErrPop
'End Sub

Sub EditInfoName()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditInfoName")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim EditRow%, EditCol%
Dim h2$

EditModus% = 1

EditRow% = flxInfo(0).row
EditCol% = flxInfo(0).col

Load frmEdit2

With frmEdit2
    .Left = picBack(0).Left + flxInfo(0).Left + flxInfo(0).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Top = picBack(0).Top + flxInfo(0).Top + EditRow% * flxInfo(0).RowHeight(0)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
    .Width = flxInfo(0).ColWidth(EditCol%)
    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit2.txtEdit
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = InfoMain.Bezeichnung(EditRow%, (EditCol% - 1) \ 2)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit2.Show 1
           
If (EditErg%) Then
    If (Trim$(EditTxt$) <> "") Then
        InfoMain.Bezeichnung(EditRow%, (EditCol% - 1) \ 2) = EditTxt$
        Call AuswahlKurzInfo
    End If
End If

Call clsError.DefErrPop
End Sub

Sub EditArbeitName()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditArbeitName")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim EditRow%, EditCol%
'Dim h2$
'
'EditModus% = 1
'
'EditRow% = 0
'EditCol% = flxarbeit(0).col
'
'Load frmEdit2
'
'With frmEdit2
'    .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(EditCol%) + 45
'    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
'    .Top = picBack(0).Top + flxarbeit(0).Top + EditRow% * flxInfo(0).RowHeight(0)
'    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
'    .Width = flxarbeit(0).ColWidth(EditCol%)
'    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
'End With
'With frmEdit2.txtEdit
'    .Width = frmEdit2.ScaleWidth
'    .Left = 0
'    .Top = 0
'    h2$ = ArbeitMain.Bezeichnung(EditCol% - Match1.OrgCols%)
'    .text = h2$
'    .BackColor = vbWhite
'    .Visible = True
'End With
'
'frmEdit2.Show 1
'
'If (EditErg%) Then
'    If (Trim$(EditTxt$) <> "") Then
'        ArbeitMain.Bezeichnung(EditCol% - Match1.OrgCols%) = EditTxt$
'        flxarbeit(0).TextMatrix(0, EditCol%) = EditTxt$
'    End If
'End If

Call clsError.DefErrPop
End Sub

Sub InitDateiButtons()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitDateiButtons")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, max%

''max% = UBound(MatchAnzeigeTyp)
'max% = Match1.UBoundMatchAnzeigeTyp
'For i% = 1 To max%
'    Load mnuDateiInd(i%)
'    Load cmdDatei(i%)
'Next i%
'
'For i% = 0 To max%
'    cmdDatei(i%).Top = 0
'    cmdDatei(i%).Left = i% * 900
'    cmdDatei(i%).Visible = True
'    cmdDatei(i%).ZOrder 1
'Next i%
'
'mnuDateiInd(0).Caption = "&Taxe"
'mnuDateiInd(1).Caption = "Taxe &phonetisch"
'mnuDateiInd(2).Caption = "&Lagerartikel"
'
'cmdDatei(0).Caption = "&T"
'cmdDatei(1).Caption = "&P"
'cmdDatei(2).Caption = "&L"

Call clsError.DefErrPop
End Sub

Sub HoleIniWerte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleIniWerte")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim l&
Dim h$
    
h$ = "N"
l& = GetPrivateProfileString(INI_SECTION, "ArtikelStatistik", "N", h$, 2, INI_DATEI)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    ArtikelStatistik% = True
Else
    ArtikelStatistik% = False
End If
    
Call clsError.DefErrPop
End Sub

Private Sub mnuZusatzinfo_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuZusatzInfo_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
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

Call clsError.DefErrPop
End Sub

Private Sub ArtikelZuordnungBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ArtikelZuordnungBefuellen")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ssatz%, sMax%, ok%, bm%, AnzLiefFuerHerst%
Dim prozent!, bmo!
Dim h$, pzn$, LiefFuerHerst$(20)
    
ZuordLief% = Val(StammdatenPzn$)

With frmLiefStammdaten.lstStammdaten3
    For i% = 0 To (.ListCount - 1)
        If (i% >= 20) Then Exit For
        .ListIndex = i%
        LiefFuerHerst$(i%) = .text
    Next i%
    AnzLiefFuerHerst% = .ListCount
End With

With frmLiefStammdaten
    ZuordAbBm% = Val(Trim(.txtStammdaten3(0).text))

'    LiefFuerHerst$(0) = UCase(Trim(.txtStammdaten3(4).text))
'    LiefFuerHerst$(1) = UCase(Trim(.txtStammdaten3(13).text))
'    LiefFuerHerst$(2) = UCase(Trim(.txtStammdaten3(14).text))
    
    ZuordBevorratungsZeit% = Val(.txtStammdaten3(9).text)
    If (ZuordBevorratungsZeit% = 0) Then ZuordBevorratungsZeit% = Val(.txtStammdaten3(6).text)
    If (ZuordBevorratungsZeit% = 0) Then ZuordBevorratungsZeit% = Para1.BestellPeriode
End With
    
picProgress.Visible = True
picProgress.SetFocus
DoEvents
flxarbeit(0).Redraw = False
KeinRowColChange% = True

Ass1.GetRecord (1)
sMax% = (Ass1.DateiLen / Ass1.RecordLen) - 1
If (sMax% <= 0) Then Call clsError.DefErrPop: Exit Sub

For ssatz% = 1 To sMax%
    Ass1.GetRecord
    
    pzn$ = Ass1.pzn
    If (Val(pzn$) <> 0) Then
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Ast1.GetRecord (FabsRecno& + 1)
            
            ok% = True
            If (Ass1.lief <> ZuordLief%) Then
                ok% = False
                h$ = UCase(Trim(Ast1.herst))
                If (h$ <> "") Then
                    For i% = 0 To (AnzLiefFuerHerst% - 1) '2
                        If (h$ = LiefFuerHerst$(i%)) Then
                            ok% = True
                            Exit For
                        End If
                    Next i%
                End If
            End If
            
            If (ok%) Then
                If (Ass1.opt <= 0) Then
                    bmo! = 0
                Else
                    bmo! = Ass1.opt
'                    bmo% = Int(Ass1.opt + 0.501)
                End If
                bm% = clsOpTool.CalcDirektBM%(ZuordBevorratungsZeit%, bmo!, Ass1.poslag)
                
                h$ = pzn$ + vbTab
                If (Ass1.lief = ZuordLief%) Then h$ = h$ + Chr$(214)
                h$ = h$ + vbTab
                h$ = h$ + vbTab + Ast1.kurz + vbTab + Ast1.meng + vbTab + Ast1.meh
                h$ = h$ + vbTab + Ast1.herst
                h$ = h$ + vbTab + Format(bm%, "0")
                h$ = h$ + vbTab + Format(Ass1.opt, "0.0")
                h$ = h$ + vbTab + Format(Ass1.poslag, "0")
                h$ = h$ + vbTab + Format(Ast1.aep, "0.00")
                h$ = h$ + vbTab + Format(Ast1.AVP, "0.00")
                h$ = h$ + vbTab + Ast1.rez
                flxarbeit(0).AddItem h$
            End If
        End If
    End If

    If (ssatz% Mod 100 = 0) Then
        prozent! = (ssatz% / sMax%) * 100!
        h$ = Format$(prozent!, "##0") + " %"
        With picProgress
            .Cls
            .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
            picProgress.Print h$
            picProgress.Line (0, 0)-((prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
        End With
        
        DoEvents
'        If (BestvorsAbbruch% = True) Then
'            Exit For
'        End If
    End If
Next ssatz%

picProgress.Visible = False

With flxarbeit(0)
    If (.Rows = 1) Then
        h$ = "XXXXXXX"
        .AddItem h$
    End If
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 1
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    
    KeinRowColChange% = True
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 1) <> "") Then
            .row = i%
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellForeColor = vbBlue
        End If
    Next i%
    KeinRowColChange% = False
    
    .FillStyle = flexFillSingle
End With

Call SortiereArtikelZuordnung

Call clsError.DefErrPop
End Sub

Private Sub SortiereArtikelZuordnung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SortiereArtikelZuordnung")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ssatz%, sMax%, ok%
Dim prozent!
Dim h$, pzn$, LiefFuerHerst$(2)
    
With flxarbeit(0)
    .Redraw = False
    
    For i% = 1 To (.Rows - 1)
        If (SortModus% = 0) Then
            h$ = ""
        ElseIf (SortModus% = 1) Then
            h$ = .TextMatrix(i%, 6)
        Else
            If (.TextMatrix(i%, 1) <> "") Then
                h$ = ""
            Else
                h$ = "*"
            End If
        End If
        .TextMatrix(i%, 2) = h$
    Next i%
    .row = 1
    .col = 2
    .RowSel = .Rows - 1
    .ColSel = 5
    .Sort = 5
    
    .TopRow = 1
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .Redraw = True
    .SetFocus
End With

Call clsError.DefErrPop
End Sub

Private Sub tmrStart_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tmrStart_Timer")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

tmrStart.Enabled = False
Call ArtikelZuordnungBefuellen

Call clsError.DefErrPop
End Sub

Private Sub picBack_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picBack_MouseMove")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Static OrgX!, OrgY!

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (iNewLine) Then
    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
        OrgX = x
        OrgY = y
        
        Call opToolbar.ShowToolbar
    
'        With nlcmd
'            nlcmd.Line (0, 0)-(.ScaleWidth, 150), RGB(75, 75, 75), BF
'            nlcmd.Line (0, 150)-(.ScaleWidth, 165), RGB(80, 80, 80), BF
'            nlcmd.Line (0, 180)-(.ScaleWidth, 195), RGB(85, 85, 85), BF
'            nlcmd.Line (0, 210)-(.ScaleWidth, 225), RGB(90, 90, 90), BF
'            Call wpara.FillGradient(nlcmd, 0 / Screen.TwipsPerPixelX, (225) / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(40, 40, 40), RGB(160, 160, 160))
'        End With
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picToolbar_MouseDown")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim start
Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, index%
Dim h$

If (iNewLine) Then
    Call opToolbar.Click(x)
'    Call opToolbar.ShowToolbar
'
'    With picToolbar
'        index = 0
'        xx = 0
'        cmdToolbarSize = 1410
'        loch = 1140 'cmdToolbarSize   '200
'        IconWidth% = 32 * Screen.TwipsPerPixelX
'        For i% = 1 To 19
'            If ((i% = 9) Or (i% = 17)) Then xx% = xx% + loch% '105
'
'    '        picToolbar.Line (x% + cmdToolbarSize%, 90)-(x% + cmdToolbarSize% + 15, .ScaleHeight - 90), RGB(150, 150, 150), BF
'            If (x >= xx) And (x <= xx% + cmdToolbarSize%) Then
'                index = i
'                Exit For
'            End If
'
'            xx% = xx% + cmdToolbarSize% + 15
'        Next i%
'        picToolbar.Line (xx, 0)-(xx% + cmdToolbarSize%, .ScaleHeight), RGB(200, 200, 200), BF
'        .PaintPicture imgToolbar(1).ListImages(index + 1).Picture, xx + 150, 150, IconWidth%, IconWidth%
'        .DrawWidth = 3
'        picToolbar.Line (xx% + 15, 15)-(xx% + cmdToolbarSize% - 15, .ScaleHeight - 15), vbWhite, B
'        .DrawWidth = 1
'    End With
'
'    DoEvents
'    start = Timer
'    Do
'        If (Timer - start) > 0.2 Then
'            Exit Do
'        End If
'    Loop
'    Call opToolbar.ShowToolbar
'
'    cmdToolbar(index).Value = True
Else
    picToolbar.Drag (vbBeginDrag)
    opToolbar.DragX = x
    opToolbar.DragY = y
End If

Call clsError.DefErrPop
End Sub

Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picToolbar_MouseMove")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, index%
Dim h$
Static OrgX!, OrgY!

If (iNewLine) Then
    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
        OrgX = x
        OrgY = y
        
'        Call opToolbar.ShowToolbar
        Call opToolbar.MouseMove(x)
    End If
End If

'If (iNewLine) Then
'    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
'        OrgX = x
'        OrgY = y
'
'        Call opToolbar.ShowToolbar
'
'        With picToolbar
'            index = 0
'            xx = 0
'            cmdToolbarSize = 1410
'            loch = 1140 'cmdToolbarSize   '200
'            For i% = 1 To 19
'                If ((i% = 9) Or (i% = 17)) Then xx% = xx% + loch% '105
'
'        '        picToolbar.Line (x% + cmdToolbarSize%, 90)-(x% + cmdToolbarSize% + 15, .ScaleHeight - 90), RGB(150, 150, 150), BF
'                If (x >= xx) And (x <= xx% + cmdToolbarSize%) Then
'                    index = i
'                    Exit For
'                End If
'
'                xx% = xx% + cmdToolbarSize% + 15
'            Next i%
'            If (cmdToolbar(index).Enabled) Then
'                .DrawWidth = 3
'                picToolbar.Line (xx% + 15, 15)-(xx% + cmdToolbarSize% - 15, .ScaleHeight - 15), vbWhite, B
'                .DrawWidth = 1
'            End If
'        End With
'    End If
'End If

Call clsError.DefErrPop
End Sub

Private Sub picBack_Paint(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picBack_Paint")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim x1&, x2&, y1&, Y2&, iAdd&, iAdd2&

If (Para1.Newline) Then
    Call wPara1.picBackPaint(picBack(0), opBereich.InfoZusatz)
'    With picBack(0)
'        .ForeColor = RGB(180, 180, 180) ' vbWhite
'        .FillStyle = vbSolid
'        .FillColor = vbWhite
'        RoundRect .hdc, (txtMatchcode.Left - 60) / Screen.TwipsPerPixelX, (txtMatchcode.Top - 30) / Screen.TwipsPerPixelY, (txtMatchcode.Left + txtMatchcode.Width + 60) / Screen.TwipsPerPixelX, (txtMatchcode.Top + txtMatchcode.Height + 15) / Screen.TwipsPerPixelY, 10, 10
'    End With
End If

Call clsError.DefErrPop
End Sub

 


