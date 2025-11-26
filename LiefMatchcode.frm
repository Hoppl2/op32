VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmLiefMatchcode 
   Caption         =   "Matchcode - Auswahl"
   ClientHeight    =   5805
   ClientLeft      =   285
   ClientTop       =   600
   ClientWidth     =   13470
   KeyPreview      =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   13470
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   11880
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3840
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
      Left            =   11760
      ScaleHeight     =   360
      ScaleWidth      =   1095
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10440
      Top             =   4680
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   10560
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin VB.ListBox lstSortierung 
      Height          =   255
      Left            =   10320
      Sorted          =   -1  'True
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDatei 
      Height          =   375
      Index           =   0
      Left            =   10560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox picSave 
      Height          =   1095
      Left            =   10440
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox picToolbar 
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
      Left            =   6600
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   11
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
         TabIndex        =   12
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   10095
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   420
         Index           =   1
         Left            =   7200
         TabIndex        =   3
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         _Version        =   393216
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         ScrollBars      =   2
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   420
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         _Version        =   393216
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   450
         Index           =   0
         Left            =   3360
         TabIndex        =   5
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1200
      End
      Begin VB.TextBox txtMatchcode 
         Appearance      =   0  '2D
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFlexBack 
         Height          =   615
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2520
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfoZusatz 
         Height          =   780
         Index           =   0
         Left            =   480
         TabIndex        =   17
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
         Left            =   210
         TabIndex        =   4
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
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
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
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   9615
      End
      Begin VB.Label lblMatchcode 
         Caption         =   "&Name/Pzn"
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
         TabIndex        =   9
         Top             =   720
         Width           =   1935
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
         TabIndex        =   8
         Top             =   270
         Width           =   9615
      End
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   2
      Left            =   12360
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":0852
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":10A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":18F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":299A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":59EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":623E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":6A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":9AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":A334
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":AB86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":DBD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":10C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1147C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":144CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":17520
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   3
      Left            =   12360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1A572
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1B2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1C016
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1CD68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1DABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":1E80C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":2185E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":225B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":23302
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":26354
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":270A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":27DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":2AE4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":2DE9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":2EBEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":31C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":34C92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   4
      Left            =   12360
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":37CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":39136
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":3A588
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":3B9DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":3CE2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":3E27E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":412D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":42722
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":43B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":46BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":48018
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":4946A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":4C4BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":4F50E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":50960
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":539B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":56A04
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
            Picture         =   "LiefMatchcode.frx":59A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":59B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":59DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":59F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5A226
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5A4B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5A74A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5A9DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5ACF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5B010
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5B122
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5B3B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5B646
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5B8D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5BB6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5BD44
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5BFD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5C0E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5C1FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5C48C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5C59E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5C8B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5CBD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5CEEC
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
            Picture         =   "LiefMatchcode.frx":5D206
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5D498
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5D7B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5DACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5DDE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5E078
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5E392
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5E624
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5E93E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5EC58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5EEEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5F17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5F40E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5F6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5F932
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5FC4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":5FEDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60170
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60402
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60694
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60926
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":60F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LiefMatchcode.frx":61274
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
         Caption         =   ""
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
         Caption         =   "Stift-Erfassung"
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
Attribute VB_Name = "frmLiefMatchcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const INI_DATEI = "\user\winop.ini"
'Const INI_SECTION = "Matchcode"
'Const INFO_SECTION = "Matchcode Infobereich"

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich
Dim ArbeitMain As clsArbeitBereich

Dim HochfahrenAktiv%

Dim Standard%

Dim ArtikelStatistik%

Dim ProgrammModus%

Dim FabsErrf%
Dim FabsRecno&

Dim NachAufruf%

Dim INI_DATEI As String

Dim AnzAnzeige%


Private Const DefErrModul = "LIEFMATCHCODE.FRM"

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
        mnuBearbeitenInd(MENU_SF7).Enabled = Dp04Ok%
        mnuBearbeitenInd(MENU_SF8).Enabled = False
        
        cmdOk(0).default = True
        cmdEsc(0).Cancel = True
        
        flxarbeit(0).SelectionMode = flexSelectionByRow
        flxarbeit(0).col = 0
        flxarbeit(0).ColSel = flxarbeit(0).Cols - 1
        
        If (iNewLine) Then
            flxarbeit(0).BackColorSel = RGB(135, 61, 52)
            flxInfo(0).BackColorSel = RGB(135, 61, 52)
        Else
            flxarbeit(0).BackColorSel = vbHighlight
            flxInfo(0).BackColorSel = vbHighlight
        End If
        
        txtMatchcode.Enabled = True
        
        h$ = Me.Caption
        ind% = InStr(h$, " (EDITIER-MODUS)")
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        Me.Caption = h$
        
    Case 1
        If (Trim(flxarbeit(0).TextMatrix(1, 1)) = "") Then
            txtMatchcode.text = "a"
            h$ = RTrim(UCase(txtMatchcode.text))
            erg% = Match2.SuchArtikel%(h$, opBereich.ArbeitAnzZeilen)
        End If
        
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

        flxarbeit(0).SelectionMode = flexSelectionFree
        flxarbeit(0).col = flxarbeit(0).Cols - 1
        flxarbeit(0).ColSel = flxarbeit(0).col
        
        flxarbeit(0).BackColorSel = vbMagenta
        flxInfo(0).BackColorSel = vbMagenta
        
        txtMatchcode.Enabled = False
               
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

Private Sub cmdDatei_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdDatei_Click")
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
Dim erg%
Dim iSuch$

Match2.MatchModus = Index
iSuch$ = Match2.OrgSuch
erg% = Match2.SuchArtikel%(iSuch$, opBereich.ArbeitAnzZeilen)

Call clsError.DefErrPop
End Sub

Private Sub cmdEsc_Click(Index As Integer)
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
With flxarbeit(1)
    If (.Rows > 1) Then
        If (Match2.MatchListeTyp) Then
            For i% = 1 To .Rows - 1
                .row = i%
                If (.CellFontItalic = False) Then
                    For j% = 4 To 7
                        If (Trim(.TextMatrix(i%, j%)) = "") Then
                            .row = i%
                            .col = j%
                            Beep
                            .SetFocus
                            Call clsError.DefErrPop: Exit Sub
                        End If
                    Next j%
                End If
            Next i%
        End If
        For i% = 1 To .Rows - 1
            .row = i%
            If (.CellFontItalic = False) Then
                Match2.MatchcodeTxt = Trim$(.TextMatrix(i%, 1)) + "  " + Trim$(.TextMatrix(i%, 2)) + " " + Trim$(.TextMatrix(i%, 3))
                h$ = h$ + .TextMatrix(i%, 0) + "@" + Match2.MatchcodeTxt + "@"
                h$ = h$ + .TextMatrix(i%, 4) + "@" + .TextMatrix(i%, 5) + "@" + .TextMatrix(i%, 6)
                If (Match2.MatchListeTyp) Then
                    h$ = h$ + "@" + .TextMatrix(i%, 7) + "@" + .TextMatrix(i%, 8) + "@" + .TextMatrix(i%, 9)
                End If
                h$ = h$ + vbTab
            End If
        Next i%
    
        If (h$ <> "") Then
            Match2.MatchcodePzn = .TextMatrix(1, 0)
        End If
    End If
End With

Match2.MatchcodeErg = h$

Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdOk_Click(Index As Integer)
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
    If (ActiveControl.Name = txtFlexBack.Name) Then
        col% = flxarbeit(0).col - Match2.OrgCols%
        If (col% >= 0) Then
            MatchTyp% = Match2.MatchTyp
            erg% = ArbeitMain.EditArbeitBelegung(col%)
            If (erg%) Then
                Call MachAuswahlGrid(False)
                Call MachMaske
            End If
        End If
    ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
        With flxInfo(0)
            row% = .row
            col% = .col
            h$ = RTrim(.text)
        End With
        If (col% Mod 2) Then
            MatchTyp% = Match2.MatchTyp
            Call InfoMain.EditInfoBelegung
            Call AuswahlKurzInfo
        End If
    End If
ElseIf (ActiveControl.Name = txtMatchcode.Name) Then
    flxarbeit(0).Visible = True
    txtFlexBack.Visible = True
    flxarbeit(1).Visible = False
    h$ = RTrim(UCase(txtMatchcode.text))
    If (Left$(h$, 1) = "#") Then h$ = Mid$(h$, 2)
    erg% = Match2.SuchArtikel%(h$, opBereich.ArbeitAnzZeilen)
ElseIf (ActiveControl.Name = flxarbeit(0).Name) Then
    If (ActiveControl.Index = 1) Then
        Call Match2.EditSatz
    End If
ElseIf (ActiveControl.Name = txtFlexBack.Name) Then
    h$ = UCase(RTrim(flxarbeit(0).TextMatrix(flxarbeit(0).row, 1)))
    ind% = InStr(h$, "SIEHE")
    If (ind% > 0) Then
        h$ = Mid$(h$, ind% + 6)
        Do
            ind% = InStr(h$, " ")
            If (ind% > 0) Then
                h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
            Else
                Exit Do
            End If
        Loop
        erg% = Match2.SuchArtikel%(h$, opBereich.ArbeitAnzZeilen)
    Else
'        If (MatchTyp% = MATCH_LIEFERANTEN) Or (MatchTyp% = MATCH_HILFSTAXE) Then
        If (Match2.MatchNurEiner) Then
            row% = flxarbeit(0).row
            Match2.MatchcodePzn = Format$(Ausgabe(row% - 1).pzn, "0000000")
            Match2.MatchcodeErg = Match2.MatchcodePzn
            Unload Me
        Else
            Call Match2.ListeAusFlexBefuellen
        End If
    End If
ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
    With flxInfo(0)
        row% = .row
        col% = .col
        h$ = RTrim(.text)
    End With
    If (col% = 0) Then
        If (h$ = TEXT_ABSAGEN) Or (h$ = TEXT_NACHBEARBEITUNG) Then
            row% = flxarbeit(0).row
            ind% = Ausgabe(row% - 1).pzn
            txt$ = flxarbeit(0).TextMatrix(row%, 1)
            Call clsDialog.AnzeigeFenster(h$, ind%, txt$)
        End If
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdToolbar_Click(Index As Integer)
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
If (Index = 0) Then
'    Me.WindowState = vbMinimized
ElseIf (Index <= 8) Then
    Call mnuBearbeitenInd_Click(Index - 1)
ElseIf (Index <= 16) Then
    Call mnuBearbeitenInd_Click(Index)
ElseIf (Index = 19) Then
'    Call mnuBeenden_Click
End If

Call clsError.DefErrPop
End Sub

Private Sub comSenden_OnComm()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("comSenden_OnComm")
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

Dim s As String
Dim i As Integer
Dim s1 As String

Select Case comSenden.CommEvent
    Case comEvReceive
        If comSenden.PortOpen Then
            s1 = comSenden.Input
            If (s1 = vbCr) Then
                cmdOk(0).Value = True
            Else
                With txtMatchcode
                    s = Trim(.text)
                    If (.SelLength > 0) Then
                        s = Mid$(s, .SelLength + 1)
                    End If
                    .text = s + s1
                End With
            End If
        End If
End Select

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_DblClick(Index As Integer)
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

Private Sub flxarbeit_GotFocus(Index As Integer)
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
If (Index = 0) Then
    txtFlexBack.SetFocus
    AuswahlKurzInfo
End If
Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_RowColChange(Index As Integer)
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

Private Sub flxInfo_DblClick(Index As Integer)
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

Private Sub flxInfo_GotFocus(Index As Integer)
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

If (Index = 0) Then
    ArbeitRow% = flxarbeit(0).row
    pzn$ = Format$(Ausgabe(ArbeitRow% - 1).pzn, "0000000")
    
    With flxInfo(0)
        arow% = .row
        aCol% = .col
        
        InfoRow% = 0
        
        .TextMatrix(InfoRow%, 0) = TEXT_ABSAGEN
        InfoRow% = InfoRow% + 1
        .TextMatrix(InfoRow%, 0) = TEXT_NACHBEARBEITUNG
        InfoRow% = InfoRow% + 1
        
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

INI_DATEI = CurDir + "\winop.ini"
   
'Top = frmAction.Top + 600
'Left = frmAction.Left + 600
'Width = frmAction.Width - 1200
'Height = frmAction.Height - 1200

'Top = 0
'Left = 0
If (iNewLine) Then
    WindowState = ProjektForm.WindowState
    Width = ProjektForm.Width
    Height = ProjektForm.Height
    Top = ProjektForm.Top
    Left = ProjektForm.Left
    Call wPara1.ControlBorderless(Me, 7, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 7)   '+ 3
Else
    Width = Screen.Width - (1200 * wPara1.BildFaktor)
    Height = Screen.Height - (1200 * wPara1.BildFaktor)
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End If


Caption = Match2.IniSection
lblMatchcode.Caption = Match2.EingabeName

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
End With



Set opToolbar = New clsToolbar
Call opToolbar.InitToolbar(Me, INI_DATEI, Match2.IniSection$, "23112412")

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
cmdToolbar(14).ToolTipText = "shift+F7 Stift-Erfassung"
cmdToolbar(15).ToolTipText = "shift+F8"
cmdToolbar(16).ToolTipText = "shift+F9"
'cmdToolbar(19).ToolTipText = "Programm beenden"


Call wPara1.InitFont(Me)
Call HoleIniWerte

Match2.MatchModus = Standard%

Set ArbeitMain = New clsArbeitBereich
Call ArbeitMain.InitArbeitBereich(flxarbeit(0), INI_DATEI, Match2.ArbeitSection$)

Set InfoMain = New clsInfoBereich
Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, Match2.InfoSection$)
Call InfoMain.ZeigeInfoBereich("", False)
Call ZeigeInfoBereichAdd(0)
flxInfo(0).row = 0
flxInfo(0).col = 0

Set opBereich = New clsOpBereiche
Call opBereich.InitBereich(Me, opToolbar)
opBereich.AutoRedraw = 0
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = True
opBereich.ArbeitWasDarunter = False
opBereich.InfoTitel = False
opBereich.InfoZusatz = ArtikelStatistik%
opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = -2

mnuZusatzInfo.Checked = ArtikelStatistik%

Call InitDateiButtons

txtMatchcode.text = ""
If (Match2.MatchcodeTxt <> "") Then
    txtMatchcode.text = Match2.MatchcodeTxt
End If

ProgrammModus% = 0

'flxarbeit(0).Rows = 2
'flxarbeit(0).Cols = 1
flxarbeit(0).row = 1
Call WechselModus(0)

flxarbeit(0).Visible = False
With flxarbeit(1)
    .Rows = 2
    .FixedRows = 1
    .Cols = 2
    .FormatString = "|<Lieferanten"
    .Rows = 1
    .Visible = True
End With

If (SeriellScannerOk%) Then
    erg% = clsOpTool.OpenCom(Me, SeriellScannerParam$)
    If (erg% = False) Then Call clsDialog.MessageBox("Seriell-Scanner nicht verfügbar!", vbExclamation)
End If

HochfahrenAktiv% = False

NachAufruf% = True

If (Match2.MatchAutoRet) Then tmrStart.Enabled = True

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
            
Call MachAuswahlGrid
'Call Match2.MachAuswahlGrid(opBereich.ArbeitAnzZeilen)
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

Sub MachAuswahlGrid(Optional AllesNeu% = True)
Dim i%, j%, spBreite%, OrgCols%, arow%, aCol%, MaxBreite%, AnzArbeitCols%, iAusr%, iKz%
Dim h$

With flxarbeit(0)
    .Redraw = False
    
    If (AllesNeu%) Then
        .Rows = opBereich.ArbeitAnzZeilen
        .FixedRows = 1

        Call Match2.InitAuswahlGrid
    End If
    
    AnzArbeitCols% = ArbeitMain.AnzArbeitSpalten
    OrgCols% = Match2.OrgCols
    .Cols = OrgCols% + AnzArbeitCols%
    
    Font.Bold = True
    For i% = 1 To AnzArbeitCols%
        j% = OrgCols% - 1 + i%
        .ColWidth(j%) = TextWidth(String(ArbeitMain.Laenge(i% - 1), "X"))
        .TextMatrix(0, j%) = ArbeitMain.Bezeichnung(i% - 1)
    Next i%
    Font.Bold = False
    
    If (AnzArbeitCols% > 0) Then
        KeinRowColChange% = True
        arow% = .row
        aCol% = .col
        For i% = 1 To AnzArbeitCols%
            .FillStyle = flexFillRepeat
            .row = 0
            .col = OrgCols% - 1 + i%
            .RowSel = .Rows - 1
            .ColSel = .col
            iAusr% = ArbeitMain.Ausrichtung(i% - 1)
            If (iAusr% = 0) Then
                .CellAlignment = flexAlignLeftCenter
            ElseIf (iAusr% = 2) Then
                .CellAlignment = flexAlignRightCenter
            Else
                .CellAlignment = flexAlignCenterCenter
            End If
            iKz% = ArbeitMain.IstKennzeichen(i% - 1)
            If (iKz%) Then
                .CellFontName = "Courier New"
            Else
                .CellFontName = wPara1.FontName(0)
            End If
            .FillStyle = flexFillSingle
        Next i%
        .row = arow%
        .col = aCol%
        KeinRowColChange% = False
    End If
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (i% <> 1) Then
'            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
            spBreite% = spBreite% + .ColWidth(i%)
        End If
    Next i%
    MaxBreite% = .Width - 60
    If (spBreite% > MaxBreite%) Then
        spBreite% = MaxBreite%
    End If
    .ColWidth(1) = MaxBreite% - spBreite%
    
    .Redraw = True
End With

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
Dim i%, spBreite%

On Error Resume Next

ReDim Preserve Ausgabe(opBereich.ArbeitAnzZeilen - 1)
AnzAnzeige% = opBereich.ArbeitAnzZeilen - 1
lblMatchcode.Left = wPara1.LinksX
lblMatchcode.Top = wPara1.TitelY   'FlexY%
txtMatchcode.Left = lblMatchcode.Left + lblMatchcode.Width + 150
txtMatchcode.Top = lblMatchcode.Top

If (iNewLine) Then
    With txtMatchcode
'        .Appearance = 0
        .BackColor = vbWhite
        Call wPara1.ControlBorderless(txtMatchcode, 1, 1)
    End With
    With picBack(0)
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtMatchcode.Left - 90) / Screen.TwipsPerPixelX, (txtMatchcode.Top - 45) / Screen.TwipsPerPixelY, (txtMatchcode.Left + txtMatchcode.Width + 90) / Screen.TwipsPerPixelX, (txtMatchcode.Top + txtMatchcode.Height + 45) / Screen.TwipsPerPixelY, 10, 10
    End With
End If



txtFlexBack.Top = flxarbeit(0).Top + 15
txtFlexBack.Left = flxarbeit(0).Left

flxarbeit(1).Left = flxarbeit(0).Left
flxarbeit(1).Top = flxarbeit(0).Top
flxarbeit(1).Width = flxarbeit(0).Width
flxarbeit(1).Height = flxarbeit(0).Height
With flxarbeit(1)
    Font.Bold = True
    
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(1) = .Width - spBreite% - 90
End With

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
Dim i%

On Error Resume Next

lblMatchcode.BackColor = wPara1.FarbeArbeit

flxarbeit(1).BackColor = vbWhite
flxarbeit(1).BackColorBkg = vbWhite

Call clsError.DefErrPop
End Sub

Sub MachMaske()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachMaske")
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
Dim i%, j%, erg%, h$, ind%, ind2%, sAnz%, NurLagerndeAktiv%
Dim OrgControl As Control

If (txtMatchcode.text = "") Then Call clsError.DefErrPop: Exit Sub

Ausgabe(0).Name = Left$(Ausgabe(0).Name, Ausgabe(0).Verweis) + ArbeitMain.ZeigeArbeitBereich(Format(Ausgabe(0).pzn, "0000000"))

AnzAnzeige% = opBereich.ArbeitAnzZeilen - 1
For i% = 1 To (AnzAnzeige% - 1)
    erg% = Match2.SuchWeiter%(i% - 1, True)
    If (erg%) Then
        Call Match2.Umspeichern(buf$, i%)
        Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis) + ArbeitMain.ZeigeArbeitBereich(Format(Ausgabe(i%).pzn, "0000000"))
    Else
        AnzAnzeige% = i%
        Exit For
    End If
Next i%

Set OrgControl = ActiveControl
Call AuswahlBefüllen
OrgControl.SetFocus
With flxarbeit(0)
    .row = 1
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Sub AuswahlBefüllen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AuswahlBefüllen")
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
Dim i%, j%, k%, ind%, AltRow%, AltCol%, AnzArbeitCols%, OrgCols%, iKz%, aFontBold%
Dim Suc&
Dim h$

On Error Resume Next

With flxarbeit(0)
    KeinRowColChange% = True

    txtFlexBack.Visible = False
    .Visible = False
    AltRow% = .row
    AltCol% = .col
    
    For i% = 1 To AnzAnzeige%
        
        h$ = Ausgabe(i% - 1).Name
        For j% = 0 To .Cols - 2
            ind% = InStr(h$, vbTab)
            .TextMatrix(i%, j%) = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Next j%
        .TextMatrix(i%, .Cols - 1) = RTrim$(h$)
    Next i%
        
    .row = 1
    .col = .Cols - 1
    h$ = .CellFontName
    
    
    .FillStyle = flexFillRepeat
    AnzArbeitCols% = ArbeitMain.AnzArbeitSpalten
    OrgCols% = Match2.OrgCols
    
    For i% = 1 To AnzAnzeige%
        .row = i%
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1

        If (Ausgabe(i% - 1).LagerKz = 2) Then
            .CellFontBold = True
        Else
            .CellFontBold = False
        End If
    Next i%
    .FillStyle = flexFillSingle
    
    For i% = 1 To AnzArbeitCols%
        iKz% = ArbeitMain.IstKennzeichen(i% - 1)
        If (iKz%) Then
            For j% = 1 To AnzAnzeige%
                .row = j%
                .col = OrgCols% - 1 + i%
                aFontBold% = .CellFontBold
                .CellFontName = "Courier New"
                .CellFontBold = aFontBold%
            Next j%
'        Else
'            .CellFontName = wPara1.FontName(0)
        End If
    Next i%
    
    
    .Rows = AnzAnzeige% + 1
    
'    If (MatchModus% = LAGER_MATCH) Then
'        For i% = 1 To AnzAnzeige%
'            .row = i%
'            .col = .Cols - 1
'            .CellFontName = "Courier New"
'        Next i%
'    End If
    
    .row = AltRow%
    .col = AltCol%
    .Visible = True
    txtFlexBack.Visible = True
    Call AuswahlKurzInfo
    KeinRowColChange% = False
End With
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

row% = flxarbeit(0).row

ActKontrollen$ = ""
actzuordnung$ = ""
ActZusatz$ = ""

If (flxarbeit(0).Visible) Then
    pzn$ = Format(Ausgabe(row% - 1).pzn, "0000000")
Else
    With flxarbeit(1)
        If (.Rows > 1) Then
            pzn$ = .TextMatrix(.row, 0)
        Else
            pzn$ = "XXXXXXX"
        End If
    End With
End If

iRow% = flxInfo(0).row
iCol% = flxInfo(0).col
Call InfoMain.ZeigeInfoBereich(pzn$, True)
Call ZeigeInfoBereichAdd(0)
flxInfo(0).row = iRow%
flxInfo(0).col = iCol%

If (opBereich.InfoZusatz) Then
    Call ZeigeInfoZusatz(pzn$)
End If

Call clsError.DefErrPop
End Sub

Sub ZeigeInfoBereichAdd(Index%)
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

With flxInfo(Index%)
    .Redraw = False
    
    .row = 0
    .col = 0
    .CellFontBold = True
    .TextMatrix(0, 0) = Match2.MatchAnzeigeTyp(Match2.MatchModus)
    
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

'With flxInfoZusatz(0)
'    .Redraw = False
'    .GridLines = flexGridInset
'    .SelectionMode = flexSelectionFree
'    .Rows = 2
'    .FixedRows = 1
'    .Cols = 15
'
'    .FillStyle = flexFillRepeat
'    .col = 0
'    .row = 1
'    .ColSel = .Cols - 1
'    .RowSel = .Rows - 1
'    .CellBackColor = vbWhite
'
'    .col = 0
'    .row = 0
'    .ColSel = .Cols - 1
'    .RowSel = .Rows - 1
'    .CellAlignment = flexAlignCenterCenter
'    .FillStyle = flexFillSingle
'
'    Set ArtStat1 = New clsArtStatistik
'    erg% = ArtStat1.StatistikRechnen(pzn$)
'    If (erg%) Then
'        AltJahr& = -1
'        j% = 0
'        For i% = 0 To 12
'            Termin& = ArtStat1.Anfang - i% - 1
'            Jahr& = (Termin& - 1) \ 12
'            Monat% = Termin& - Jahr& * 12
'            If (AltJahr& <> Jahr&) Then
'                .TextMatrix(0, j%) = Str$(Jahr&)
'                iWert! = ArtStat1.JahresWert(Jahr&)
'                If (iWert! <> 0!) Then
'                    .TextMatrix(1, j%) = Str$(iWert!)
'                Else
'                    .TextMatrix(1, j%) = ""
'                End If
'                .row = 0
'                .col = j%
'                .CellFontBold = True
'                .row = 1
'                .CellFontBold = True
'                j% = j% + 1
'                AltJahr& = Jahr&
'            End If
'
'            .TextMatrix(0, j%) = Para1.MonatKurz(Monat%)
'
'            iWert! = ArtStat1.MonatsWert(i% + 1)
'            If (iWert! = 0) Then
'                .TextMatrix(1, j%) = ""
'            Else
'                .TextMatrix(1, j%) = Str$(iWert!)
'            End If
'            j% = j% + 1
'        Next i%
'    Else
'        For i% = 0 To .Cols - 1
'            .TextMatrix(0, i%) = ""
'            .TextMatrix(1, i%) = ""
'        Next i%
'    End If
'    Set ArtStat1 = Nothing
'
'    .Redraw = True
'End With


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
If (comSenden.PortOpen) Then comSenden.PortOpen = False
Call clsError.DefErrPop
End Sub

Private Sub mnuBearbeitenInd_Click(Index As Integer)
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
Dim i%, erg%, row%, col%, ind%, iAusr%
Dim l&
Dim pzn$, txt$

If (iNewLine) Then
    ind = Index
    If (ind <= MENU_F9) Then
        ind = ind + 1
    End If
    Call opToolbar.Click(-ind)
End If

Select Case Index
    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = txtFlexBack.Name) Then
                col% = flxarbeit(0).col - Match2.OrgCols%
                If (col% >= -1) Then
                    Call ArbeitMain.InsertArbeitBelegung(col% + 1)
                    Call ArbeitMain.EditArbeitBelegung(col% + 1)
                    If (ArbeitMain.Bezeichnung(col% + 1) = "") Then
                        Call ArbeitMain.LoescheArbeitBelegung(col% + 1)
                    Else
                        Call MachAuswahlGrid(False)
                        flxarbeit(0).col = col% + Match2.OrgCols% + 1
                        Call MachMaske
                    End If
                End If
            ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call MachMaske
                Call AuswahlKurzInfo
            End If
        Else
            txtMatchcode.text = ""
            flxarbeit(1).Visible = True
            flxarbeit(0).Visible = False
            txtFlexBack.Visible = False
            txtMatchcode.SetFocus
        End If
    
    Case MENU_F3
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = txtFlexBack.Name) Then
                col% = flxarbeit(0).col - Match2.OrgCols%
                If (col% >= 0) Then
                    iAusr% = ArbeitMain.Ausrichtung(col%)
                    iAusr% = (iAusr% + 1) Mod 3
                    ArbeitMain.Ausrichtung(col%) = iAusr%
                    Call MachAuswahlGrid(False)
                    Call MachMaske
                End If
            End If
        Else
'            erg% = clsDialog.WechselFenster(MatchAnzeigeTyp$, Standard%)
'            l& = WritePrivateProfileString(Match1.IniSection$, "Standard", Str$(Standard%), INI_DATEI)
'            If (erg% >= 0) Then
'                cmdDatei(erg%).Value = True
'            End If
            erg% = Match2.MatchModus + 1
            If (erg% > Match2.UBoundMatchAnzeigeTyp$) Then erg% = 0
'            If (erg% > UBound(MatchAnzeigeTyp$)) Then erg% = 0
            cmdDatei(erg%).Value = True
        End If
        
    Case MENU_F5
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = txtFlexBack.Name) Then
                col% = flxarbeit(0).col - Match2.OrgCols%
                If (col% >= 0) Then
                    Call ArbeitMain.LoescheArbeitBelegung(col%)
                    Call MachAuswahlGrid(False)
                    Call MachMaske
                End If
            ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call MachMaske
                Call AuswahlKurzInfo
            End If
        ElseIf (flxarbeit(1).Visible) Then
            With flxarbeit(1)
                .Redraw = False
                If (.Rows > 1) Then
                    If (.TextMatrix(.row, 7) = "*") Then
                        .TextMatrix(.row, 7) = ""
                        iAusr% = False
                    Else
                        .TextMatrix(.row, 7) = "*"
                        iAusr% = True
                    End If
                    col% = .col
                    For i% = 0 To (.Cols - 1)
                        .col = i%
                        .CellFontItalic = iAusr%
                    Next i%
                    .col = col%
                End If
                .Redraw = True
            End With
        End If
        
    Case MENU_F8
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = txtFlexBack.Name) Then
                col% = flxarbeit(0).col - Match2.OrgCols%
                If (col% >= 0) Then
                    Call EditArbeitName
                End If
            ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
                col% = flxInfo(0).col
                If (col% Mod 2) Then
                    row% = flxInfo(0).row
                    If (InfoMain.Bezeichnung(row%, (col% - 1) \ 2) <> "") Then
                        Call EditInfoName
                    End If
                End If
            End If
        ElseIf (ActiveControl.Name = txtMatchcode.Name) Then
        Else
            With flxarbeit(0)
                If (RTrim$(.TextMatrix(.row, 1)) <> "") Then
                    ind% = Ausgabe(.row - 1).pzn
                    txt$ = RTrim$(.TextMatrix(.row, 1))
                    Call clsDialog.ZusatzFenster(ZUSATZ_LIEFERANTEN, ind%, txt$)
                End If
            End With
        End If

    Case MENU_SF3
        row% = flxarbeit(0).row
        If (row% > 0) Then
            txt$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
            If (Val(txt$) > 0) Then Call clsDialog.Stammdaten(txt$)
        End If
'    Case MENU_SF3
'        If (ActiveControl.Name = txtFlexBack.Name) Then
'            If (MatchTyp% = MATCH_LIEFERANTEN) Then
'                row% = flxarbeit(0).row
'                Call clsDialog.Stammdaten(Format$(Ausgabe(row% - 1).pzn, "0000000"))
'            End If
'        End If
        
    Case MENU_SF4
        row% = flxarbeit(0).row
        If (row% > 0) Then
            txt$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
            If (Val(txt$) > 0) Then
                StammdatenPzn$ = txt$
                frmInternet.Show 1
            End If
        End If
'    Case MENU_SF4
'        If (ActiveControl.Name = txtFlexBack.Name) Then
'            If (MatchTyp% = MATCH_LIEFERANTEN) Then
'                row% = flxarbeit(0).row
'                StammdatenPzn$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
'                frmInternet.Show 1
'            End If
'        End If
        
    
    Case MENU_SF7
        Call Dp04Einlesen
        
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

Private Sub mnuDateiInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("mnuDateiInd_Click")
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
cmdDatei(Index).Value = True
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
cmdOk(0).Value = True

Call clsError.DefErrPop
End Sub

Private Sub txtFlexBack_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtFlexBack_GotFocus")
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
flxarbeit(0).HighLight = flexHighlightAlways
Call clsError.DefErrPop
End Sub

Private Sub txtFlexBack_lostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtFlexBack_lostFocus")
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
flxarbeit(0).HighLight = flexHighlightNever
Call clsError.DefErrPop
End Sub

Private Sub txtFlexBack_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtFlexBack_KeyDown")
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
Dim erg%, col%, iLaenge%

Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
        Call AuswahlRowChange(KeyCode)
        KeyCode = 0
    Case vbKeyLeft
        If (ProgrammModus% = 1) Then
            If (flxarbeit(0).col > 0) Then
                flxarbeit(0).col = flxarbeit(0).col - 1
            End If
        End If
        KeyCode = 0
    Case vbKeyRight
        If (ProgrammModus% = 1) Then
            If (flxarbeit(0).col < flxarbeit(0).Cols - 1) Then
                flxarbeit(0).col = flxarbeit(0).col + 1
            End If
        End If
        KeyCode = 0
'    Case 107, 187
'        If (ProgrammModus% = 1) Then
'            col% = flxarbeit(0).col - Match1.OrgCols%
'            If (col% >= 0) Then
'                ArbeitMain.Laenge(col%) = ArbeitMain.Laenge(col%) + 1
'                Call MachAuswahlGrid(False)
'            End If
'        End If
'    Case 109, 189
'        If (ProgrammModus% = 1) Then
'            col% = flxarbeit(0).col - Match1.OrgCols%
'            If (col% >= 0) Then
'                iLaenge% = ArbeitMain.Laenge(col%)
'                If (iLaenge% > 1) Then
'                    ArbeitMain.Laenge(col%) = iLaenge% - 1
'                    Call MachAuswahlGrid(False)
'                End If
'            End If
'        End If
End Select

If (ProgrammModus% = 0) Then
    flxarbeit(0).ColSel = flxarbeit(0).Cols - 1
End If

Call clsError.DefErrPop
End Sub

Private Sub txtFlexBack_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtFlexBack_KeyPress")
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
Dim col%, iLaenge%

If (ProgrammModus% = 1) Then
    col% = flxarbeit(0).col - Match2.OrgCols%
    If (col% >= 0) Then
        Select Case KeyAscii
            Case Asc("+")
                ArbeitMain.Laenge(col%) = ArbeitMain.Laenge(col%) + 1
                Call MachAuswahlGrid(False)
            Case Asc("-")
                iLaenge% = ArbeitMain.Laenge(col%)
                If (iLaenge% > 1) Then
                    ArbeitMain.Laenge(col%) = iLaenge% - 1
                    Call MachAuswahlGrid(False)
                End If
        End Select
    End If
End If

Call clsError.DefErrPop

End Sub

Sub AuswahlRowChange(KeyCode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AuswahlRowChange")
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
Dim erg%, i%, j%, h$, ind%, neu%, NurLagerndeAktiv%, sAnz%
    
neu% = True
With flxarbeit(0)
    Select Case KeyCode
        Case vbKeyUp
            If (.row > 1) Then
                .row = .row - 1
                neu% = False
            Else
                erg% = Match2.SuchWeiter%(0, False)
                If (erg%) Then
                    GoSub ZeileHinein
                Else
                    .row = 1
                End If
            End If
        Case vbKeyPageUp
            erg% = Match2.SuchWeiter%(0, False)
            If (erg%) Then
                For j% = 1 To (opBereich.ArbeitAnzZeilen - 1)
                    GoSub ZeileHinein
                    erg% = Match2.SuchWeiter%(0, False)
                    If (erg% = False) Then Exit For
                Next j%
            Else
                .row = 1
            End If
        Case vbKeyDown
            If (.row < .Rows - 1) Then
                .row = .row + 1
                neu% = False
            Else
                erg% = Match2.SuchWeiter%(AnzAnzeige% - 1, True)
                If (erg%) Then
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    i% = AnzAnzeige% - 1
                    Call Match2.Umspeichern(buf$, i%)
                    Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis) + ArbeitMain.ZeigeArbeitBereich(Format(Ausgabe(i%).pzn, "0000000"))
                Else
                    .row = AnzAnzeige%
                End If
            End If
        Case vbKeyPageDown
            erg% = Match2.SuchWeiter%(AnzAnzeige% - 1, True)
            If (erg%) Then
                For j% = 1 To AnzAnzeige%
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    i% = AnzAnzeige% - 1
                    Call Match2.Umspeichern(buf$, i%)
                    Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis) + ArbeitMain.ZeigeArbeitBereich(Format(Ausgabe(i%).pzn, "0000000"))
                    erg% = Match2.SuchWeiter%(AnzAnzeige% - 1, True)
                    If (erg% = False) Then Exit For
                Next j%
            Else
                .row = AnzAnzeige%
            End If
    End Select
End With


If (neu% = True) Then Call AuswahlBefüllen
txtFlexBack.SetFocus
Call clsError.DefErrPop: Exit Sub

ZeileHinein:
For i% = (opBereich.ArbeitAnzZeilen - 2) To 1 Step -1
    Ausgabe(i%) = Ausgabe(i% - 1)
Next i%
Call Match2.Umspeichern(buf$, 0)
Ausgabe(0).Name = Left$(Ausgabe(0).Name, Ausgabe(0).Verweis) + ArbeitMain.ZeigeArbeitBereich(Format(Ausgabe(0).pzn, "0000000"))
If (AnzAnzeige% < (opBereich.ArbeitAnzZeilen - 1)) Then
    AnzAnzeige% = AnzAnzeige% + 1
    flxarbeit(0).Rows = AnzAnzeige% + 1
End If
Return

Call clsError.DefErrPop
End Sub

Private Sub txtMatchcode_GotFocus()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtMatchcode_GotFocus")
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
With txtMatchcode
    .SelStart = 0
    .SelLength = Len(.text)
End With

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

Private Sub mnuToolbarPositionInd_Click(Index As Integer)
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

opToolbar.Position = Index

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
Call MachMaske
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
Dim erg%
Dim h$

On Error Resume Next

If (HochfahrenAktiv%) Then Call clsError.DefErrPop: Exit Sub

If (Me.WindowState = vbMinimized) Then Call clsError.DefErrPop: Exit Sub

Call opBereich.ResizeWindow

If (NachAufruf%) Then
    flxarbeit(0).Visible = True
    flxarbeit(0).Rows = flxarbeit(0).FixedRows
    txtFlexBack.Visible = True
    flxarbeit(1).Visible = False
    txtMatchcode.SetFocus
End If
If (txtMatchcode.text <> "") Then
    If (NachAufruf%) Then
'        cmdOk(0).Value = True
    Else
        Call MachMaske
    End If
End If

'If (NachAufruf%) Then txtMatchcode.SetFocus

NachAufruf% = False

picSave.Visible = False

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(flxarbeit(Index), picBack(Index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub flxInfo_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(flxInfo(Index), picBack(Index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(lblArbeit(Index), picBack(Index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub lblInfo_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(lblInfo(Index), picBack(Index), Source, x, y)
Call clsError.DefErrPop
End Sub

Private Sub picBack_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(picBack(Index), picBack(Index), Source, x, y)
Call clsError.DefErrPop
End Sub

'Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
'opToolbar.DragY = y
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
Dim EditRow%, EditCol%
Dim h2$

EditModus% = 1

EditRow% = 0
EditCol% = flxarbeit(0).col

Load frmEdit2

With frmEdit2
    .Left = picBack(0).Left + flxarbeit(0).Left + flxarbeit(0).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Top = picBack(0).Top + flxarbeit(0).Top + EditRow% * flxInfo(0).RowHeight(0)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
    .Width = flxarbeit(0).ColWidth(EditCol%)
    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit2.txtEdit
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = ArbeitMain.Bezeichnung(EditCol% - Match2.OrgCols%)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit2.Show 1
           
If (EditErg%) Then
    If (Trim$(EditTxt$) <> "") Then
        ArbeitMain.Bezeichnung(EditCol% - Match2.OrgCols%) = EditTxt$
        flxarbeit(0).TextMatrix(0, EditCol%) = EditTxt$
    End If
End If

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

'max% = UBound(MatchAnzeigeTyp)
max% = Match2.UBoundMatchAnzeigeTyp
For i% = 1 To max%
    Load mnuDateiInd(i%)
    Load cmdDatei(i%)
Next i%

For i% = 0 To max%
    cmdDatei(i%).Top = 0
    cmdDatei(i%).Left = i% * 900
    cmdDatei(i%).Visible = True
    cmdDatei(i%).ZOrder 1
Next i%

mnuDateiInd(0).Caption = "&Lieferanten"
mnuDateiInd(1).Caption = "Lieferanten &numerisch"
cmdDatei(0).Caption = "&L"
cmdDatei(1).Caption = "&N"

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
Dim i%, ind%, iVal%
Dim l&, f&
Dim h$, h2$, key$
    
With frmLiefMatchcode
    
    h$ = "0"
    l& = GetPrivateProfileString(Match2.IniSection, "Standard", "0", h$, 2, INI_DATEI)
    Standard% = Val(Left$(h$, l&))
    
    h$ = "N"
    l& = GetPrivateProfileString(Match2.IniSection, "ArtikelStatistik", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ArtikelStatistik% = True
    Else
        ArtikelStatistik% = False
    End If
    
End With
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
l& = WritePrivateProfileString(Match2.IniSection, "ArtikelStatistik", h$, INI_DATEI)

opBereich.RefreshBereich
Call MachMaske

Call clsError.DefErrPop
End Sub

Sub Dp04Einlesen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Dp04Einlesen")
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
Dim Dp04Deb%, Dp04Menge%
Dim h$
        
flxarbeit(1).Visible = True
flxarbeit(0).Visible = False
txtFlexBack.Visible = False
txtMatchcode.SetFocus

Call clsDialog.LeseStiftEinlesen
'frmDp04.Show 1
Dp04Deb% = clsDat.FileOpen("winwawi.dp0", "I")
Do While Not (EOF(Dp04Deb%))
    Line Input #Dp04Deb%, h$
    If (Left$(h$, 1) = "<") Then
        h$ = clsOpTool.PruefeDp04Zeile(h$, Dp04Menge%)
        Call Match2.ListeAusStiftBefuellen(h$, Dp04Menge%)
'        Call SucheZeile(True, Dp04Menge%)
    End If
Loop
Close #Dp04Deb%

Call clsError.DefErrPop
End Sub

Sub GetDDE(cmdStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetDDE")
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

txtMatchcode.text = Left$(cmdStr$, Len(cmdStr$) - 1)
DoEvents
cmdOk(0).Value = True

Call clsError.DefErrPop
End Sub



'Private Sub picBack_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("picBack_MouseMove")
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
'Static OrgX!, OrgY!
'
'If (picToolTip.Visible = True) Then
'    picToolTip.Visible = False
'End If
'
'If (iNewLine) Then
'    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
'        OrgX = x
'        OrgY = y
'
'        Call opToolbar.ShowToolbar
'
''        With nlcmd
''            nlcmd.Line (0, 0)-(.ScaleWidth, 150), RGB(75, 75, 75), BF
''            nlcmd.Line (0, 150)-(.ScaleWidth, 165), RGB(80, 80, 80), BF
''            nlcmd.Line (0, 180)-(.ScaleWidth, 195), RGB(85, 85, 85), BF
''            nlcmd.Line (0, 210)-(.ScaleWidth, 225), RGB(90, 90, 90), BF
''            Call wpara.FillGradient(nlcmd, 0 / Screen.TwipsPerPixelX, (225) / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(40, 40, 40), RGB(160, 160, 160))
''        End With
'    End If
'End If
'
'Call clsError.DefErrPop
'End Sub
'
'Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
'Dim start
'Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, index%
'Dim h$
'
'If (iNewLine) Then
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
'Else
'    picToolbar.Drag (vbBeginDrag)
'    opToolbar.DragX = x
'    opToolbar.DragY = y
'End If
'
'Call clsError.DefErrPop
'End Sub
'
'Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("picToolbar_MouseMove")
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
'Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, index%
'Dim h$
'Static OrgX!, OrgY!
'
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
'
'Call clsError.DefErrPop
'End Sub
 
Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub flxarbeit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxarbeit_MouseMove")
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

If (iNewLine) Then
'    Call opToolbar.ShowToolbar
    Call opToolbar.Click(x)
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
Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, Index%
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

Call clsError.DefErrPop
End Sub

Private Sub picBack_Paint(Index As Integer)
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
    With picBack(0)
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtMatchcode.Left - 60) / Screen.TwipsPerPixelX, (txtMatchcode.Top - 30) / Screen.TwipsPerPixelY, (txtMatchcode.Left + txtMatchcode.Width + 60) / Screen.TwipsPerPixelX, (txtMatchcode.Top + txtMatchcode.Height + 15) / Screen.TwipsPerPixelY, 10, 10
    End With
End If

Call clsError.DefErrPop
End Sub


