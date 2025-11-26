VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAction 
   Caption         =   "Bestellung"
   ClientHeight    =   7890
   ClientLeft      =   60
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
   Icon            =   "winwawi.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   7890
   ScaleWidth      =   11295
   Begin VB.ListBox lstDirektSortierung 
      Height          =   300
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrOptimal 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   10680
      Top             =   4920
   End
   Begin VB.Timer tmrRowa 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
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
      Left            =   10920
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDatei 
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   19
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
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstLieferant 
      Height          =   300
      Left            =   840
      TabIndex        =   16
      Top             =   120
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
      TabIndex        =   15
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
         TabIndex        =   12
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.TextBox txtDDEServer 
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Begin MSFlexGridLib.MSFlexGrid flxEinzelBewertung 
         Height          =   735
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1296
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorFixed  =   -2147483624
         HighLight       =   0
         GridLines       =   0
         ScrollBars      =   0
      End
      Begin VB.CommandButton cmdEsc 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   450
         Index           =   0
         Left            =   5280
         TabIndex        =   17
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
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1200
      End
      Begin VB.PictureBox picBestellWerte 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   1320
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   3960
         Index           =   0
         Left            =   240
         TabIndex        =   8
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
      Begin VB.Timer tmrAction 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   720
         Top             =   1080
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   1500
         Index           =   0
         Left            =   600
         TabIndex        =   10
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
         TabIndex        =   22
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
         TabIndex        =   9
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
         TabIndex        =   11
         Top             =   270
         Width           =   9615
      End
   End
   Begin VB.CommandButton cmdAltR 
      Caption         =   "&R"
      Height          =   375
      Left            =   10920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid flxSortierung 
      Height          =   975
      Left            =   10200
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   4
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxDirektAufteilung 
      Height          =   735
      Left            =   10800
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1296
      _Version        =   393216
      BackColor       =   -2147483624
      BackColorFixed  =   -2147483624
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   0
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
            Picture         =   "winwawi.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":09EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":0D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":1022
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":12B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":15CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":18E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":1F1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":21AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":24C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":27E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":2A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":2D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":30A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":33C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":36DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":396E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":3C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":3E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":41AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":44C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":47E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":4AFA
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
            Picture         =   "winwawi.frx":4E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":4F26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":51B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":52CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":55E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":5876
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":5B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":5E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":613C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":6456
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":66E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":6A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":6D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":6FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":72C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":75E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":78FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":7C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":7D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":7E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":7F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":8266
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":8580
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":889A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winwawi.frx":8BB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Beste&llung"
         Index           =   0
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "&Warenübernahme"
         Index           =   2
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Besorger-&Verwaltung"
         Index           =   3
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
         Caption         =   "S&onderangebote"
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
         Caption         =   "Wa&renwert"
         Index           =   14
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Bestell&vorschlag"
         Index           =   15
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Rück&kauf-Anfrage"
         Index           =   16
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mnuDummy11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBearbeitenLayout 
         Caption         =   "La&yout editieren"
      End
      Begin VB.Menu mnuDummy13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRueckrufe 
         Caption         =   "Erwünschte Rückrufe"
         Begin VB.Menu mnuRueckrufeInd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDummy14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiefStammdaten 
         Caption         =   "Lieferanten-Stammdaten"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPosPartner 
         Caption         =   "Lagerstand Partner-Apos"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEditAufteilung 
         Caption         =   "Aufteilung Direktbezug bearbeiten"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDummy17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAepKalkulation 
         Caption         =   "Aep-Kalkulation"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuTeilAbschluss 
         Caption         =   "Teil-Abschluss"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDummy15 
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
      Begin VB.Menu mnuZusatzInfo 
         Caption         =   "Artikel-S&tatistik"
      End
      Begin VB.Menu mnuEinzelBewertung 
         Caption         =   "&Einzel-Bewertung"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDirektAufteilung 
         Caption         =   "Aufteilung Direktbezug"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDummy10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLetztGesendeter 
         Caption         =   "Letzter Sende&vorgang"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuGesendet 
         Caption         =   "&Letzte Sendevorgänge"
         Begin VB.Menu mnuGesendetInd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuSchwellwert 
         Caption         =   "Schwell&wert-Automatik"
         Begin VB.Menu mnuSchwellwertInd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDirektbezug 
         Caption         =   "&Direktbezugs-Automatik"
         Begin VB.Menu mnuDirektbezugInd 
            Caption         =   ""
            Index           =   0
         End
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

'Const INI_DATEI = "\user\winop.ini"

Const INI_SECTION = "Bestellung"
Const INFO_SECTION = "Infobereich Bestellung"


'Dim scrAuswahlAltValue%
Dim InRowColChange%

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

Dim HochfahrenAktiv%

Dim Standard%

Dim LetztGesendete$(9)
Dim LetztSchwellwerte$(9)
Dim LetztDirektbezuege$(9)

Private Const DefErrModul = "WINWAWI.FRM"

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
Dim i%, ind%
Dim h$

Select Case ProgrammModus%
'    Case ZEILEN
'        picBack(0).Visible = False
'        cmdOk(0).Default = False
'        cmdEsc(0).Cancel = False
'        tmrAction.Enabled = False
'        AnzeigeLaufNr& = 0&
'    Case BESTSENDEN
'        If (seriell% And (comSenden.PortOpen)) Then comSenden.PortOpen = False
    
End Select

'If (Bereiche(NeuerModus%).BereichOk = False) Then
'    Call InitBereich(Me, Bereiche(NeuerModus%), NeuerModus%)
'End If
        
Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        mnuExtras.Enabled = True
        
        mnuBearbeitenInd(MENU_F2).Enabled = True
        mnuBearbeitenInd(MENU_F3).Enabled = True
        mnuBearbeitenInd(MENU_F4).Enabled = True
        mnuBearbeitenInd(MENU_F5).Enabled = True
        mnuBearbeitenInd(MENU_F6).Enabled = True 'false
        mnuBearbeitenInd(MENU_F7).Enabled = True
        mnuBearbeitenInd(MENU_F8).Enabled = True
        mnuBearbeitenInd(MENU_F9).Enabled = True
        mnuBearbeitenInd(MENU_SF2).Enabled = True
        mnuBearbeitenInd(MENU_SF3).Enabled = (BestellAnzeige% = 2)  'True
        mnuBearbeitenInd(MENU_SF4).Enabled = True
        mnuBearbeitenInd(MENU_SF5).Enabled = True
        mnuBearbeitenInd(MENU_SF6).Enabled = ModemOk% And (BestellAnzeige% = 2)
        mnuBearbeitenInd(MENU_SF7).Enabled = True
        mnuBearbeitenInd(MENU_SF8).Enabled = BestVorsAktiv%
        
        mnuBearbeitenLayout.Checked = False
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        flxarbeit(0).BackColorSel = vbHighlight
        flxInfo(0).BackColorSel = vbHighlight
        
        If (ProgrammChar$ = "B") Then tmrAction.Enabled = True
        
        h$ = Me.Caption
        ind% = InStr(h$, " (EDITIER-MODUS)")
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        Me.Caption = h$
    Case 1
        mnuDatei.Enabled = False
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = False
        mnuExtras.Enabled = False
        
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

Private Sub aniSuchen_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("aniSuchen_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

Private Sub cmdAltR_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAltR_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Call ActProgram.AltRClick
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

picQuittieren.Visible = False
ProgrammTyp% = Index
Call InitProgrammTyp
HochfahrenAktiv% = False
Call Form_Resize
Call InitProgramm
If (picQuittieren.Visible = False) Or (vAnzeigeSperren% = False) Or (ProgrammChar$ = "W") Then flxarbeit(0).SetFocus

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

If (Index = 0) Then
    If (Shift And 2) Then
        If (KeyCode = vbKeyV) Then
            EingabeStr$ = Clipboard.GetText
        ElseIf (KeyCode = vbKeyInsert) Or (KeyCode = vbKeyC) Then
            Clipboard.SetText flxarbeit(Index).Clip, vbCFText
            KeyCode = 0
        End If
    End If
End If

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
    Call ActProgram.flxArbeitKeyPress(KeyAscii)
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
EingabeStr$ = ""
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
        
        Call ActProgram.flxInfoGotFocus(InfoRow%)
        
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
EingabeStr$ = ""

If (KeinRowColChange% = False) Then
'    Error 6
    Call EchtKurzInfo
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

EingabeStr$ = ""

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (Index = 0) Then
'    If ((ProgrammChar$ = "B") And (flxarbeit(0).redraw = True) And (KeinRowColChange% = False)) Then
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

On Error Resume Next

HochfahrenAktiv% = True
   
prot% = True
If (prot% = True) Then
    PROTOKOLL% = FreeFile
    Open "winwawi.fnt" For Output As PROTOKOLL%
End If



Call wpara.InitEndSub(Me)
Set opToolbar = New clsToolbar

Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI)
Call wpara.InitFont(Me)
Call HoleIniWerte

'ProgrammTyp% = Standard%


Set InfoMain = New clsInfoBereich
Set opBereich = New clsOpBereiche


If (prot% = True) Then
'    Close #PROTOKOLL%  '????
End If
    
Set SendeForm = frmSenden

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

tmrRowa.Enabled = False

If (comSenden.PortOpen) Then comSenden.PortOpen = False

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
    .Visible = True
End With

h$ = ProgrammNamen$(ProgrammTyp%)
Caption = h$ + " - "
ProgrammChar$ = Left$(h$, 1)

On Error Resume Next
For i% = 1 To 5
    Unload mnuBearbeitenZusatzInd(i%)
Next i%
On Error GoTo DefErr

If (ProgrammTyp% = 0) Then
    Set ActProgram = New clsBestellung
    flxarbeit(0).SelectionMode = flexSelectionFree
ElseIf (ProgrammTyp% = 2) Then
    Set ActProgram = New clsWarenÜber
    flxarbeit(0).SelectionMode = flexSelectionFree
ElseIf (ProgrammTyp% = 3) Then
    Set ActProgram = New clsBesorger
'    flxarbeit(0).SelectionMode = flexSelectionByRow
End If

cmdAltR.Enabled = False

Call ActProgram.Init(Me, opToolbar, InfoMain, opBereich)
ErstAuslesen% = True

picBack(0).Visible = True

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
If (comSenden.PortOpen) Then comSenden.PortOpen = False
Call Programmende
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
Dim erg%, row%, col%
Dim l&
Dim h$, mErg$

Select Case Index

    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call EchtKurzInfo
            End If
        ElseIf (ProgrammChar$ <> "V") Then
            tmrAction.Enabled = False
            mErg$ = ArtikelAuswahl$
            If (mErg$ <> "") Then Call ActProgram.ManuellFertig
            If (ProgrammChar$ = "B") Then tmrAction.Enabled = True
        End If
    
    Case MENU_F3
        If (ProgrammChar$ = "B") Then
            BestellAnzeige% = (BestellAnzeige% + 1) Mod 3
            Standard% = BestellAnzeige%
            l& = WritePrivateProfileString(INI_SECTION, "Standard", Str$(Standard%), INI_DATEI)
            BekartCounter% = -1
            With tmrOptimal
                .Enabled = False
                .Interval = 300
                .Enabled = True
            End With
'            Call ActProgram.AuslesenBestellung(True, False, True)
        ElseIf (ProgrammChar$ = "V") Then
            BestellAnzeige% = (BestellAnzeige% + 1) Mod 3
            Call ActProgram.AuslesenBesorger
        ElseIf (IstAltLast%) Then
            h$ = WuLifDat$
            If (Len(h$) = 2) And (Mid$(h$, 2, 1) <> " ") Then
                erg% = WechselFenster(AltLastNamen$, Standard%)
                If (erg% >= 0) Then
                    Mid$(h$, 2, 1) = Mid$(AltLastStr$, erg% + 1, 1)
                    WuLifDat$ = h$
                    Call ActProgram.AuslesenWu
                End If
            End If
        End If
    
    Case MENU_F4
        If (ProgrammChar$ <> "V") Then
            Call SetKommentarTyp(1)
        End If
        
    Case MENU_F5
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call EchtKurzInfo
            End If
        Else
            Call ActProgram.MenuBearbeiten(Index)
        End If
        
    Case MENU_F6
        If (ProgrammChar$ = "B") Then
            Call DruckeBestellung
        ElseIf (ProgrammChar$ = "W") Then
            Call DruckeWu
        End If
    
    Case MENU_F7
        Call ActProgram.MenuBearbeiten(Index)
        
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
        Else
            Call ActProgram.MenuBearbeiten(Index)
        End If
    
    Case MENU_SF2
        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF3
        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF4
        Call ActProgram.MenuBearbeiten(Index)
            
    Case MENU_SF5
        Call ZeigeStatistik
        
    Case MENU_SF6
        Call ActProgram.MenuBearbeiten(Index)
        
    Case MENU_SF7
        Call ActProgram.MenuBearbeiten(Index)

    Case MENU_SF8
        Call ActProgram.MenuBearbeiten(Index)

    Case MENU_SF9
        Call ActProgram.MenuBearbeiten(Index)
'        If (ProgrammModus% = 0) Then
'            Call WechselModus(1)
'        Else
'            Call WechselModus(0)
'        End If
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

Private Sub mnuDirektbezugInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDirektbezugInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

h$ = LetztDirektbezuege$(Index)
If (Trim(h$) <> "") Then
    GesendetDatei$ = h$
    frmDirektProtokoll.Show 1
End If

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
    If (Index < 2) Then
        Call opBereich.ResizeWindow
    ElseIf (Index = 2) Then
        Call ActProgram.RefreshDunklerBereich
    End If
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
    frmAction.flxInfo(0).Clear
    Call ActProgram.mnuFontClick
End If

Call DefErrPop
End Sub

Private Sub mnuGesendetInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuGesendetInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

h$ = LetztGesendete$(Index)
If (Trim(h$) <> "") Then
    GesendetDatei$ = h$
    frmGesendet.Show 1
End If

Call DefErrPop
End Sub

Private Sub mnuLetztGesendeter_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuLetztGesendeter_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Call mnuGesendetInd_Click(0)
Call DefErrPop
End Sub

Private Sub mnuLiefStammdaten_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuLiefStammdaten_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

h$ = "02000"
If (ProgrammTyp% = 0) Then
    If ((BestellAnzeige% = 2) And (IstDirektLief%)) Then h$ = "02091"
    Call Stammdaten(Format(Lieferant%, "000"), Val(h$), DirektBewertung#, ActProgram)
    Call ActProgram.MenuBearbeiten(MENU_F7)
Else
    With flxarbeit(0)
        Call Stammdaten(Format(Asc(Left$(.TextMatrix(.row, 16), 1)), "000"), Val(h$))
    End With
End If

Call DefErrPop
End Sub

Private Sub mnuPosPartner_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuPosPartner_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim iLief%
                
iLief% = ZeigePosPartner(KorrPzn$, KorrTxt$, (ProgrammChar$ = "B"))
If (iLief% > 0) Then
    Call ActProgram.MakePartnerLief(iLief%)
End If

Call DefErrPop
End Sub

Private Sub mnuEditAufteilung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuPosPartner_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim iLief%
                
frmBmPartner.Show 1
'iLief% = ZeigePosPartner(KorrPzn$, KorrTxt$, (ProgrammChar$ = "B"))
'If (iLief% > 0) Then
'    Call ActProgram.MakePartnerLief(iLief%)
'End If

Call DefErrPop
End Sub


Private Sub mnuSchwellwertInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuSchwellwertInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

h$ = LetztSchwellwerte$(Index)
If (Trim(h$) <> "") Then
    GesendetDatei$ = h$
    frmSchwellProtokoll.Show 1
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

'Private Sub mnuAktZeile_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("mnuAktZeile_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim l&
'
'With mnuAktZeile
'    If (.Checked) Then
'        .Checked = False
'    Else
'        .Checked = True
'    End If
'
'    wpara.FarbeAktZeile = Abs(.Checked)
'    Call HighlightZeile
'End With
'
'Call DefErrPop
'End Sub

Private Sub mnuRueckrufe_Click()
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
Dim i%, ind%, iLief%
Dim l&
Dim SollRR$, h$

h$ = Space$(100)
l& = GetPrivateProfileString(INI_SECTION, "Rueckrufe", " ", h$, 101, INI_DATEI)
SollRR$ = Trim(Left$(h$, l&))
    
On Error Resume Next
For i% = 0 To 9
    mnuRueckrufeInd(i%).Checked = False
    h$ = mnuRueckrufeInd(i%).Caption
    
    ind% = InStr(h$, "(")
    If (ind% > 0) Then
        h$ = Mid$(h$, ind% + 1)
        ind% = InStr(h$, ")")
        h$ = Left$(h$, ind% - 1)
        iLief% = Val(h$)
        h$ = Format(iLief%, "000")
        ind% = InStr(SollRR$, h$)
        If (ind% > 0) Then mnuRueckrufeInd(i%).Checked = True
    End If
Next i%
On Error GoTo DefErr

Call DefErrPop
End Sub

Private Sub mnuRueckrufeInd_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuRueckrufeInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ind%, iLief%
Dim l&
Dim h$, SollRR$

With mnuRueckrufeInd(Index)
    If (.Checked) Then
        .Checked = False
    Else
        .Checked = True
    End If
End With

On Error Resume Next
SollRR$ = ""
For i% = 0 To 9
    If (mnuRueckrufeInd(i%).Checked) Then
        h$ = mnuRueckrufeInd(i%).Caption
        ind% = InStr(h$, "(")
        If (ind% > 0) Then
            h$ = Mid$(h$, ind% + 1)
            ind% = InStr(h$, ")")
            h$ = Left$(h$, ind% - 1)
            iLief% = Val(h$)
            h$ = Format(iLief%, "000")
            SollRR$ = SollRR$ + h$ + ","
        End If
    End If
Next i%
On Error GoTo DefErr

l& = WritePrivateProfileString(INI_SECTION, "Rueckrufe", SollRR$, INI_DATEI)

Call ActProgram.MenuBearbeiten(MENU_F7)

Call DefErrPop
End Sub

Private Sub mnuAepKalkulation_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuAepKalkulation_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Call ActProgram.WuAepKalk
Call DefErrPop
End Sub

Private Sub mnuTeilAbschluss_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuTeilAbschluss_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Call ActProgram.TeilAbschluss
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

Private Sub mnuEinzelBewertung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuEinzelBewertung_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (mnuEinzelBewertung.Checked) Then
    mnuEinzelBewertung.Checked = False
Else
    mnuEinzelBewertung.Checked = True
End If

Call EchtKurzInfo

Call DefErrPop
End Sub

Private Sub mnuDirektAufteilung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDirektAufteilung_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (mnuDirektAufteilung.Checked) Then
    mnuDirektAufteilung.Checked = False
    flxDirektAufteilung.Visible = False
Else
    mnuDirektAufteilung.Checked = True
End If

Call EchtKurzInfo

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

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If


If ((Shift And vbShiftMask) = vbShiftMask) And ((Shift And vbAltMask) = vbAltMask) And (KeyCode = 191) Then
    ResetWbestk2Senden
End If


If (Shift And vbCtrlMask And (KeyCode <> 17)) Then
    ind% = 0
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
'        Case vbKeyS
'            ind% = -1
'        Case vbKeyF11
'            ind% = 9
    End Select
    If ((Shift And vbShiftMask) And (ind% > 0)) Then
        ind% = ind% + 8
    End If
    If (ind% > 0) Then
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

Private Sub mnuAllesMarkieren_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuAllesMarkieren_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
    .col = 0
    .row = 1
    .ColSel = .Cols - 1
    .RowSel = .Rows - 1
    .SetFocus
End With

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
Call Programmende
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
frmAbout.Show 1

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
Call ActProgram.mnuOptionenClick
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
Call PruefeObUpdate(False)
'Call AuslesenBestellung(True, False)
'tmrAction.Enabled = True

Call DefErrPop
End Sub

Private Sub tmrOptimal_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrRowa_Timer")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
            
' dieser Timer ist eigentlich für Schwellwerte in wbestk2 vorgesehen, wird hier zum verzögerten
' Start des Einlesen der Bestellung verwendet
tmrOptimal.Enabled = False
Call ActProgram.AuslesenBestellung(True, False, True)

Call DefErrPop
End Sub

Private Sub tmrrowa_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrRowa_Timer")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ScrAktForm$

ScrAktForm$ = ""
On Error Resume Next
ScrAktForm$ = Screen.ActiveForm.Name
On Error GoTo DefErr

tmrRowa.Enabled = False
If (ProgrammChar = "W") And (RowaAktiv% Or ShuttleAktiv%) Then
    If (Len(WuLifDat$) > 3) And (ScrAktForm$ = Me.Name) Then
        Call ActProgram.RowaWu
    End If
    tmrRowa.Enabled = True
End If

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
Dim row%, col%, NeueRow%
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
            Call EchtKurzInfo
        End If
    End If
ElseIf (ActiveControl.Name = flxarbeit(0).Name) Then
'    Debug.Print "OK"
    If (EingabeStr$ <> "") Then
        Call ActProgram.SucheZeile
        EingabeStr$ = ""
    Else
        Call ActProgram.EditSatz
        If (ProgrammChar$ = "W") Then
            With flxarbeit(0)
                NeueRow% = 0
                If (.col = 8) Then
                    If (NachManuellerLM% = 0) Then
                        .col = 11
                        While (para.AblK = " ") And (Left$(.TextMatrix(.row, 15), 1) = "A") And (RTrim(.TextMatrix(.row, 11)) = "")
                            .col = 11
                            Call ActProgram.EditSatz
                            NeueRow% = True
                        Wend
                    ElseIf (NachManuellerLM% = 1) Then
                        NeueRow% = True
                    End If
                ElseIf (.col = 11) Then
                    NeueRow% = True
                End If
                If (NeueRow%) Then
                    If (.row < (.Rows - 1)) Then
                        .row = .row + 1
                        .col = 8
                        If (.row >= (.TopRow + opBereich.ArbeitAnzZeilen - 2)) Then
                            .TopRow = .row - opBereich.ArbeitAnzZeilen + 2
                        End If
                    End If
                End If
            End With
        End If
    End If
ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
    With flxInfo(0)
        row% = .row
        col% = .col
        h$ = RTrim(.text)
    End With
    If (col% = 0) Then
        Call ActProgram.cmdOkClick(h$)
    End If
End If

Call DefErrPop
End Sub

Private Sub comSenden_OnComm()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("comSenden_OnComm")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

Dim s As String
Dim i As Integer
Dim s1 As String
Dim kAsc As Integer

Select Case comSenden.CommEvent
    Case comEvReceive
        If comSenden.PortOpen Then
            s1 = comSenden.Input
            If (s1 = vbCr) Then
                cmdOk(0).Value = True
            Else
                kAsc = Asc(s1)
                If ((kAsc >= 48) And (kAsc <= 57)) Then EingabeStr$ = EingabeStr$ + s1
            End If
        End If
End Select

Call DefErrPop
End Sub

Sub FieldOutput(strTemp As String, fldTemp As Field)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FieldOutput")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
    ' Report function for FieldX.

    Dim prpLoop As Property

    Debug.Print "Gültige Field-Eigenschaft in " & strTemp

    ' Durchlaufen der Properties-Auflistung des
    ' übergebenen Field-Objekts.
    For Each prpLoop In fldTemp.Properties
        ' Einige Eigenschaften sind in bestimmten

' Zusammenhängen unzulässig(z.B. die
        ' Value-Eigenschaft in einer Fields-
        ' Auflistung eines TableDef-Objekts). Versuche,
        ' eine unzulässige Eigenschaft zu verwenden,
        ' lösen Fehler aus.
        On Error Resume Next
        Debug.Print "    " & prpLoop.Name & " = " & _
            prpLoop.Value
        On Error GoTo 0
    Next prpLoop

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

Call ActProgram.EchtKurzInfo

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
Dim i%, j%, Monat%, erg%, row%, col%
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

'Public Sub OpwMenuBeenden()
'
'On Error Resume Next
'frmAction.txtCliCdUpdate.LinkTopic = "Optipharm Menü|CdUpdate"
'frmAction.txtCliCdUpdate.LinkItem = "txtSvrCdUpdate"
'frmAction.txtCliCdUpdate.LinkMode = vbLinkAutomatic
'frmAction.txtCliCdUpdate.LinkExecute "Close"
'On Error GoTo 0
'
'End Sub

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
Dim KalkAvp#, RundAvp#
Dim BekLaufNr&
Dim h$, KalkText$, DirektWerte$
Static aBekLaufNr&, aFlexRow%, aDirektWerte$

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
        
        aBekLaufNr& = -1&
        aFlexRow% = .row
        aDirektWerte$ = ""
        .HighLight = flexHighlightWithFocus
        KeinRowColChange% = False
    Else
        If (ProgrammChar$ = "B") Then
            BekLaufNr& = Val(.TextMatrix(.row, 21))
        ElseIf (ProgrammChar$ = "W") Then
            BekLaufNr& = Val(.TextMatrix(.row, 20))
        Else
            BekLaufNr& = 1
            aBekLaufNr& = 2
        End If
        If (BekLaufNr& <> aBekLaufNr&) Or (aFlexRow% <> .row) Or ((IstDirektLief%) And (.col = 8)) Then
            
            If (ProgrammChar$ = "B") And (aBekLaufNr& <> -1) And (aFlexRow% > 0) And (IstDirektLief%) Then
                DirektWerte$ = Trim$(.TextMatrix(aFlexRow%, 5)) + vbTab + Trim$(.TextMatrix(aFlexRow%, 6)) + vbTab + Trim$(.TextMatrix(aFlexRow%, 7)) + vbTab + Trim$(.TextMatrix(aFlexRow%, 26))
                If (DirektWerte$ <> aDirektWerte$) Then
                    Call ActProgram.CheckDirektAngebot(aFlexRow%)
                End If
            End If
            
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
'            If (ProgrammChar$ <> "V") Then
                .col = aCol%
'            End If
            
            Call EchtKurzInfo
            aBekLaufNr& = BekLaufNr&
            aFlexRow% = .row
            If (ProgrammChar$ <> "V") Then
                aDirektWerte$ = Trim$(.TextMatrix(.row, 5)) + vbTab + Trim$(.TextMatrix(.row, 6)) + vbTab + Trim$(.TextMatrix(.row, 7)) + vbTab + Trim$(.TextMatrix(.row, 26))
            End If
            .HighLight = flexHighlightWithFocus
            KeinRowColChange% = False
            
            ZeilenWechsel% = True
        End If
        
'        If (.col = 9) And (Left$(.TextMatrix(.row, 9), 6) = "Absage") Then
        
        If (ProgrammChar$ = "B") Then
            If (.col = 8) Then
                h$ = ToolTipSchwellwert$(Val(.TextMatrix(.row, 27)))
                If (h$ <> "") Then
                    picQuittieren.Visible = False
                    Call AnzeigeKommentar(h$, 7, 2)
                End If
            ElseIf (.col = 10) Then
                h$ = ToolTipAbsagen$(.TextMatrix(.row, 0))
                If (h$ <> "") Then
                    picQuittieren.Visible = False
                    Call AnzeigeKommentar(h$, 9, 1)
                End If
            ElseIf (mnuDatei.Enabled) And (ZeilenWechsel% = False) Then
                picQuittieren.Visible = False
            End If
        ElseIf (ProgrammChar$ = "W") Then
            picQuittieren.Visible = False
            
            If (RTrim$(.TextMatrix(.row, 0)) = "XXXXXXX") Then Call DefErrPop: Exit Sub
            
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                h$ = ""
                If (.TextMatrix(.row, 1) = "$") Then
                    h$ = ActProgram.PruefeAutomaticPreis(RundAvp#)
                End If
                If (h$ <> "") Then
                    Call AnzeigeKommentar(h$, 2, 2)
                Else
                    h$ = Trim(ww.WuText)
                    If (h$ <> "") Then Call AnzeigeKommentar(h$, 2, 2)
                End If
            End If
        End If
    End If
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
        Call HighlightZeile(True)
        .row = gef%
        If (ProgrammChar$ = "V") Then
            .col = 2
        Else
            .col = 8
        End If
        
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
        Call HighlightZeile
        Call EchtKurzInfo
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
Call ActProgram.RefreshBereichsFlexSpalten
Call DefErrPop
End Sub

Sub PruefeObUpdate(BereitsGelockt%)
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

If (ProgrammChar$ = "B") Then
'    Call CheckBestzusa
    
'    Call CheckBekart
    
    If (BereitsGelockt% = False) Then Call ww.SatzLock(1)
    
    ww.GetRecord (1)
    If (GlobBekMax% <> ww.erstmax) Or (BekartCounter% <> ww.erstcounter) Then
        ret% = True
    End If
    
    If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)
    
    Call opToolbar.UpdateFlag(ret%)

    tmrAction.Enabled = True
End If


Call DefErrPop
End Sub

Sub InitAuslesenBestellung(AnzeigeCol%, AltLief%, Max%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitAuslesenBestellung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim row%
Dim erg$, pzn$, txt$

KeinRowColChange% = True
flxarbeit(0).redraw = False

AnzeigeCol% = flxarbeit(0).col
If (AnzeigeLaufNr& = 0&) Then
    With flxarbeit(0)
        If (.Rows > 1) Then
            row% = .row
            AnzeigeLaufNr& = Val(RTrim$(.TextMatrix(row%, 21)))
        End If
    End With
End If
Call HighlightZeile(True)
KeinRowColChange% = True

If (Lieferant% = 0) Then
    Lieferant% = AltLief%
    If (Lieferant% = 0) Then
        pzn$ = ""
        txt$ = ""
        erg$ = MatchCode(1, pzn$, txt$, False, False)
        If (erg$ <> "") Then
            frmLieferantenWahl.Show 1
            If (LiefWechselOk%) Then
                Lieferant% = Val(pzn$)
            End If
        End If
    End If
End If

If (Max% > 20) Then
    Call StartAnimation(Me, "Bestellung wird eingelesen ...")
End If

Call DefErrPop
End Sub

Sub ExitAuslesenBestellung(AnzeigeCol%, Max%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ExitAuslesenBestellung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, MitBestVors%, spBreite%

With flxarbeit(0)
    .row = 1
    If (AnzeigeLaufNr& > 0&) Then
        For i% = 1 To .Rows - 1
            If (Val(.TextMatrix(i%, 21)) = AnzeigeLaufNr&) Then
                .row = i%
                Exit For
            End If
        Next i%
        AnzeigeLaufNr& = 0&
    End If

    If (.row < .TopRow) Then
        .TopRow = .row
    Else
        If (.row >= (.TopRow + opBereich.ArbeitAnzZeilen - 2)) Then
            .TopRow = .row - opBereich.ArbeitAnzZeilen + 2
        End If
'            While (.row - .TopRow >= opBereich.ArbeitAnzZeilen - 1)
'                .TopRow = .TopRow + 1
'            Wend
    End If

    .col = AnzeigeCol%
    If (.col = 0) Then
        .col = 5
    End If
    If (ErstAuslesen%) Then
        .col = 5
        ErstAuslesen% = False
    End If
    .ColSel = .col


    flxarbeit(0).redraw = True
    KeinRowColChange% = False
    
    Call HighlightZeile
    
    Call EchtKurzInfo
End With

If (BestellAnzeige% = 2) Then
    mnuBearbeitenInd(MENU_SF3).Enabled = True
    mnuBearbeitenInd(MENU_SF6).Enabled = ModemOk%
'    mnuBearbeitenZusatzInd(0).Enabled = True
    mnuBearbeitenZusatzInd(1).Enabled = True
Else
    mnuBearbeitenInd(MENU_SF3).Enabled = False
    mnuBearbeitenInd(MENU_SF6).Enabled = False
'    mnuBearbeitenZusatzInd(0).Enabled = False
    mnuBearbeitenZusatzInd(1).Enabled = False
End If
cmdToolbar(10).Enabled = mnuBearbeitenInd(MENU_SF3).Enabled
cmdToolbar(13).Enabled = mnuBearbeitenInd(MENU_SF6).Enabled

If (BestellAnzeige% = 2) And (IstDirektLief%) Then
    mnuBearbeitenZusatzInd(2).Enabled = True
    mnuBearbeitenZusatzInd(3).Enabled = True
Else
    mnuBearbeitenZusatzInd(2).Enabled = False
    mnuBearbeitenZusatzInd(3).Enabled = False
End If

BestVorsAktiv% = False
If (DarfBestVors%) And (para.BestVorsAktiv Or ((BestellAnzeige% = 2) And IstDirektLief%)) Then
    BestVorsAktiv% = True
End If
mnuBearbeitenInd(MENU_SF8).Enabled = BestVorsAktiv%
cmdToolbar(15).Enabled = mnuBearbeitenInd(MENU_SF8).Enabled

Font.Bold = True
spBreite% = TextWidth("99.99")
Font.Bold = False
spBreite% = spBreite% + TextWidth("X")
With flxarbeit(0)
    If (IstDirektLief%) Then
        If (.ColWidth(7) = 0) Then .ColWidth(2) = .ColWidth(2) - spBreite%
        .ColWidth(7) = spBreite%
    Else
        If (.ColWidth(7) > 0) Then .ColWidth(2) = .ColWidth(2) + spBreite%
        .ColWidth(7) = 0
    End If
End With

If (Max% > 20) Then
    Call StopAnimation(Me)
End If

Call opToolbar.UpdateFlag(False)

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

With frmAction
    .picBestellWerte.Font.Name = .flxarbeit(0).Font.Name
    .picBestellWerte.Font.Size = .flxarbeit(0).Font.Size
    .picBestellWerte.Top = .flxarbeit(0).Top + .flxarbeit(0).Height
    .picBestellWerte.Left = .flxarbeit(0).Left
    .picBestellWerte.Height = opBereich.ZeilenHoeheY + 90
    .picBestellWerte.Width = .flxarbeit(0).Width
    If (ProgrammChar$ <> "V") Then
        .picBestellWerte.Visible = True
        Call ActProgram.ZeigeWerte
    End If
    
    .lstLieferant.Font.Name = .flxarbeit(0).Font.Name
    .lstLieferant.Font.Size = .flxarbeit(0).Font.Size
    .lstLieferant.Top = .picBack(0).Top + .flxarbeit(0).Top
    .lstLieferant.Height = .flxarbeit(0).Height
    .lstLieferant.Left = .picBack(0).Left + .flxarbeit(0).Left + .flxarbeit(0).ColPos(7)
    .lstLieferant.Width = .TextWidth("XXXXXXXXXXXXXX")
End With
  
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
    
    
    h$ = "0"
    l& = GetPrivateProfileString(INI_SECTION, "Standard", "0", h$, 2, INI_DATEI)
    Standard% = Val(Left$(h$, l&))
    BestellAnzeige% = Standard%
    
    ProgrammTyp% = 0
    If (Command <> "") Then ProgrammTyp% = Val(Command)
    
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

    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "ArtikelStatistik", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ArtikelStatistik% = True
    Else
        ArtikelStatistik% = False
    End If
    
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
    l& = GetPrivateProfileString(INI_SECTION, "LagerKontrollListe", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        DruckeLagerKontrollListe% = True
    Else
        DruckeLagerKontrollListe% = False
    End If
    
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "AngebotY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    iVal% = Val(h$)
    If (iVal% <= 0) Then
        iVal% = -1
    End If
    AngebotY% = iVal%
        
'    h$ = Space$(8)
'    l& = GetPrivateProfileString(UserSection$, "FarbeGray", h$, h$, 9, INI_DATEI)
'    h$ = Left$(h$, l&)
'    If (Trim(h$) = "") Then
'        FarbeGray& = vbGrayText
'    Else
'        FarbeGray& = wpara.BerechneFarbWert(h$)
'    End If

    
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "BestVorsProtokoll", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        BvProtAktiv% = True
    Else
        BvProtAktiv% = False
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
    
    h$ = "00"
    l& = GetPrivateProfileString(INI_SECTION, "TageSpeichern", "00", h$, 3, INI_DATEI)
    h$ = Left$(h$, l&)
    TageSpeichern% = Val(h$)
    
    h$ = Space$(100)
    l& = GetPrivateProfileString("Allgemein", "Rowa", " ", h$, 101, "\user\dp.ini")
    If (Trim(Left$(h$, l&)) <> "") Then
        RowaAktiv% = True
    Else
        RowaAktiv% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString("Allgemein", "Shuttle", "N", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ShuttleAktiv% = True
    Else
        ShuttleAktiv% = False
    End If
    
    h$ = "J"
    l& = GetPrivateProfileString(INI_SECTION, "BesorgerSperren", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        vAnzeigeSperren% = True
    Else
        vAnzeigeSperren% = False
    End If
  
    h$ = "1"
    l& = GetPrivateProfileString(INI_SECTION, "AnzahlRetourenDruck", "1", h$, 2, INI_DATEI)
    AnzRetourenDruck% = Val(Left$(h$, l&))
    
    
    
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
    
    h$ = "3"
    l& = GetPrivateProfileString(INI_SECTION, "HandShake", "3", h$, 2, INI_DATEI)
    PharmaBoxHandShake% = Val(Left$(h$, l&))
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "PharmaboxInDOS", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ModemInDOS% = True
    Else
        ModemInDOS% = False
    End If
    
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "KalkOhnePreis", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        iVal% = True
    Else
        iVal% = False
    End If
    KalkOhnePreis% = iVal%
    
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

    
    
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "AepKalkX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    iVal% = Val(h$)
    If (iVal% < 0) Then iVal% = 0
    AepKalkX% = iVal%
        
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "AepKalkY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    iVal% = Val(h$)
    If (iVal% < 0) Then iVal% = 0
    AepKalkY% = iVal%
        
    h$ = Space$(100)
    l& = GetPrivateProfileString(INI_SECTION, "AllgemeineAngebote", h$, h$, 101, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
'    If (h$ = "") Then h$ = Printer.DeviceName
    GhAllgAngebote$ = h$

    h$ = "0"
    l& = GetPrivateProfileString(INI_SECTION, "NachManuellerLM", "0", h$, 2, INI_DATEI)
    NachManuellerLM% = Val(Left$(h$, l&))
    
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
    l& = GetPrivateProfileString(INI_SECTION, "Wbestk2ManuellSenden", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        Wbestk2ManuellSenden% = True
    Else
        Wbestk2ManuellSenden% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "Bm0Anzeigen", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        Bm0Anzeigen% = True
    Else
        Bm0Anzeigen% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "PartnerTeilBestellungen", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        PartnerTeilBestellungen% = True
    Else
        PartnerTeilBestellungen% = False
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "PartnerMitLagerstand", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        PartnerBestaendeBeruecksichtigen% = True
    Else
        PartnerBestaendeBeruecksichtigen% = False
    End If
    
    h$ = Space$(30)
    l& = GetPrivateProfileString("OpPartner", "IdBeiPartnern", h$, h$, 31, INI_DATEI)
    IdBeiPartnern$ = Left$(h$, l&)

    h$ = Space$(50)
    l& = GetPrivateProfileString(INI_SECTION, "AbsagenMitNL", h$, h$, 51, INI_DATEI)
    AbsagenMitNL$ = Left$(h$, l&)

    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "KalkNichtRezPflichtigeAM", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        KalkNichtRezPflichtigeAM% = True
    Else
        KalkNichtRezPflichtigeAM% = False
    End If
    
    h$ = Space$(50)
    l& = GetPrivateProfileString(INI_SECTION, "AutomatenLieferanten", h$, h$, 51, INI_DATEI)
    AutomatenLiefs$ = Trim(Left$(h$, l&))
    If (AutomatenLiefs$ <> "") Then
        AutomatenLac$ = Left$(h$, 1)
        AutomatenLiefs$ = Mid$(AutomatenLiefs$, 2)
    End If

    h$ = Space$(30)
    l& = GetPrivateProfileString(INI_SECTION, "LetztBesorgerWeg", h$, h$, 31, INI_DATEI)
    LetztBesorgerWeg$ = Trim(Left$(h$, l&))
    If (LetztBesorgerWeg$ = "") Then
        LetztBesorgerWeg$ = Format("311202")
    End If

    h$ = "N"
    l& = GetPrivateProfileString(INI_SECTION, "LieferantenAbfrage", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        LieferantenAbfrage% = True
    Else
        LieferantenAbfrage% = False
    End If
    
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

l& = WritePrivateProfileString(INI_SECTION, "MinutenWarnung", Str$(AnzMinutenWarnung%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "MinutenWarten", Str$(AnzMinutenWarten%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "MinutenVerspaetung", Str$(AnzMinutenVerspaetung%), INI_DATEI)

h$ = "N"
If (BestVorsKomplett%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "BestVorsKomplett", h$, INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "BestVorsKomplettMinuten", Str$(BestVorsKomplettMinuten%), INI_DATEI)

h$ = "N"
If (BestVorsPeriodisch%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "BestVorsPeriodisch", h$, INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "BestVorsPeriodischMinuten", Str$(BestVorsPeriodischMinuten%), INI_DATEI)

h$ = "N"
If (MacheEtiketten%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "Etiketten", h$, INI_DATEI)
    
h$ = "N"
If (DruckeLagerKontrollListe%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "LagerKontrollListe", h$, INI_DATEI)

h$ = "N"
If (BvProtAktiv%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "BestVorsProtokoll", h$, INI_DATEI)

h$ = "N"
If (TeilDefekte%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "TeilDefekte", h$, INI_DATEI)
    
l& = WritePrivateProfileString(INI_SECTION, "TageSpeichern", Str$(TageSpeichern%), INI_DATEI)

h$ = "N"
If (vAnzeigeSperren%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "BesorgerSperren", h$, INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "AnzahlRetourenDruck", Str$(AnzRetourenDruck%), INI_DATEI)
 
h$ = "N"
If (SchwellwertAktiv%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertAktiv", h$, INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "SchwellwertMinuten", Str$(SchwellwertMinuten%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertSicherheit", Str$(SchwellwertSicherheit%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertToleranz", Str$(SchwellwertToleranz%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertWarnungProz", Str$(SchwellwertWarnungProz%), INI_DATEI)
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertWarnungAb", Str$(SchwellwertWarnungAb%), INI_DATEI)

h$ = "N"
If (SchwellwertVorab%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertVorab", h$, INI_DATEI)

h$ = "N"
If (SchwellwertGlaetten%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "SchwellwertGlaetten", h$, INI_DATEI)

h$ = "N"
If (ModemInDOS%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(INI_SECTION, "PharmaboxInDOS", h$, INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "AutomatikDrucker", AutomatikDrucker$, INI_DATEI)

l& = WritePrivateProfileString("DirektBezug", "MinimalSendefenster", Str$(DirektBezugSendMinuten%), INI_DATEI)
l& = WritePrivateProfileString("DirektBezug", "MaxDauerKontrollen", Str$(DirektBezugKontrollenMinunten%), INI_DATEI)
l& = WritePrivateProfileString("DirektBezug", "AlleMinutenWarnung", Str$(DirektBezugWarnungMinuten%), INI_DATEI)

l& = WritePrivateProfileString(INI_SECTION, "NachManuellerLM", Str$(NachManuellerLM%), INI_DATEI)
    
h$ = "N"
If (MitLagerstandCalc%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "MitLagerstand", h$, INI_DATEI)

h$ = "N"
If (WuPruefung%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "WuPruefung", h$, INI_DATEI)

h$ = "N"
If (Wbestk2ManuellSenden%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "Wbestk2ManuellSenden", h$, INI_DATEI)

h$ = "N"
If (Bm0Anzeigen%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "Bm0Anzeigen", h$, INI_DATEI)

h$ = "N"
If (PartnerTeilBestellungen%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "PartnerTeilBestellungen", h$, INI_DATEI)

h$ = "N"
If (PartnerBestaendeBeruecksichtigen%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "PartnerMitLagerstand", h$, INI_DATEI)

h$ = "N"
If (KalkNichtRezPflichtigeAM%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "KalkNichtRezPflichtigeAM", h$, INI_DATEI)

h$ = "N"
If (LieferantenAbfrage%) Then h$ = "J"
l& = WritePrivateProfileString(INI_SECTION, "LieferantenAbfrage", h$, INI_DATEI)

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
        Call EchtKurzInfo
    End If
End If

Call DefErrPop
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

For i% = 1 To 3
    Load cmdDatei(i%)
Next i%

For i% = 0 To 3
    cmdDatei(i%).Top = 0
    cmdDatei(i%).Left = i% * 900
    cmdDatei(i%).Visible = True
    cmdDatei(i%).ZOrder 1
Next i%

cmdDatei(0).Caption = "&L"
cmdDatei(1).Caption = "&K"
cmdDatei(2).Caption = "&W"
cmdDatei(3).Caption = "&V"

Call DefErrPop
End Sub

Function InternerKommentar$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InternerKommentar$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim row%, AbholNr%, Belegt%, BlockNummer%
Dim pzn$, ch$, h$
Dim Kiste1 As clsKiste

InternerKommentar$ = ""

With flxarbeit(0)
    row% = .row
    pzn$ = .TextMatrix(row%, 0)
    pzn$ = KorrPzn$(pzn$)
    
    h$ = RTrim$(.TextMatrix(row%, 12))
    AbholNr% = Val(h$)
End With

If (AbholNr% < 0) Or (AbholNr% > 999) Then
    Call DefErrPop: Exit Function
End If

Set Kiste1 = New clsKiste

Kiste1.OpenDatei
Belegt% = Kiste1.Belegt(AbholNr%)
If (Belegt%) Then
    BlockNummer% = Kiste1.PznInKiste(AbholNr%, pzn$)
    If (BlockNummer% >= 0) Then
        Kiste1.GetInhalt (BlockNummer%)
        h$ = RTrim$(Kiste1.InfoText(3))
        If (h$ <> "") Then
            InternerKommentar$ = h$ + Chr$(13) + RTrim$(Kiste1.InfoText(4)) + Chr$(13)
        End If
    End If
End If

Kiste1.CloseDatei

Call DefErrPop
End Function

Function ToolTipAbsagen$(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToolTipAbsagen$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ind%, babMax%, babAnz%
Dim h$, h2$, ret$

ret$ = ""

Me.MousePointer = vbHourglass

absagen.OpenDatei
absagen.GetRecord (1)
babMax% = absagen.erstmax

babAnz% = (absagen.DateiLen / absagen.RecordLen) - 1

For i% = 1 To babAnz% '2000
    babMax% = babMax% + 1
    If (babMax% > babAnz%) Then babMax% = 1   '2000
    absagen.GetRecord (babMax% + 1) 'früher: i%
    If (absagen.pzn = pzn$) Then
        h2$ = CVDatum2(absagen.datum)
        h$ = Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2) + "@"
        
        lif.GetRecord (absagen.Lief + 1)
        h2$ = lif.kurz
        h2$ = h2$ + " (" + Mid$(Str$(absagen.Lief), 2) + ")"
        
        h$ = h$ + h2$ + "@"
        h$ = h$ + Str$(absagen.menge) + "@"
        h2$ = absagen.rest
        ind% = InStr(h2$, Chr$(0))
        If (ind% > 0) Then h2$ = Left$(h2$, ind% - 1)
        h$ = h$ + h2$ + "@" + Chr$(13)
        ret$ = h$ + ret$
    End If
Next i%

absagen.CloseDatei

ToolTipAbsagen$ = ret$

Me.MousePointer = vbDefault

Call DefErrPop
End Function

Function ToolTipSchwellwert$(schwell%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToolTipSchwellwert$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ret$

ret$ = ""
Select Case schwell%
    Case 1
        ret$ = "Erreichen Mindest-Umsatz"
    Case 2
        ret$ = "Günstigster Lieferant"
    Case 3
        ret$ = "Zeitfenster abgelaufen"
    Case 4
        ret$ = "Schwellwert-Sprung"
End Select

ToolTipSchwellwert$ = ret$

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
If (modus% = 0) And (vAnzeigeSperren%) Then sperren% = True

If (sperren%) Then
    ReDim Preserve KommentarOk&(UBound(KommentarOk&) + 1)
    KommentarOk&(UBound(KommentarOk)) = lNr&
    Call SetKommentarTyp(0)
End If


Call DefErrPop
End Sub

Sub SetKommentarTyp(typ%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetKommentarTyp")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, val1%

If (typ% = 0) Then
    val1% = False
    picQuittieren.SetFocus
Else
    val1% = True
    picQuittieren.Visible = False
End If

flxarbeit(0).Enabled = val1%
flxInfo(0).Enabled = val1%
picBack(0).Enabled = val1%
    
mnuDatei.Enabled = val1%
On Error Resume Next
For i% = 0 To 3
    cmdDatei(i%).Enabled = val1%
Next i%
For i% = 0 To 16
    mnuBearbeitenInd(i%).Enabled = val1%
Next i%
On Error GoTo DefErr

mnuAnsicht.Enabled = val1%
mnuExtras.Enabled = val1%
For i% = 0 To 19
    cmdToolbar(i%).Enabled = val1%
Next i%

If (typ% = 0) Then
    mnuBearbeitenInd(2).Enabled = True
    cmdToolbar(3).Enabled = True
Else
    KeinRowColChange% = True
    flxarbeit(0).SetFocus
    DoEvents
    KeinRowColChange% = False
End If

Call DefErrPop
End Sub

Sub ResetKommentar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResetKommentar")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

picQuittieren.Visible = False
flxarbeit(0).Enabled = True
flxInfo(0).Enabled = True
picBack(0).Enabled = True

mnuDatei.Enabled = True
On Error Resume Next
For i% = 0 To 16
    mnuBearbeitenInd(i%).Enabled = True
Next i%
On Error GoTo 0
mnuAnsicht.Enabled = True
mnuExtras.Enabled = True
For i% = 0 To 19
    cmdToolbar(i%).Enabled = True
Next i%

flxarbeit(0).SetFocus

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

Sub ErzeugeGesendeteMenu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeGesendeteMenu")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, iLief%, iRufzeit%
Dim h$, h2$

On Error Resume Next
For i% = 1 To 9
    Unload mnuGesendetInd(i%)
Next i%
On Error GoTo DefErr


j% = 0
For i% = 9 To 1 Step -1
    h$ = Dir("winw\*.sp" + Format(i%, "0"))
    If (h$ <> "") Then
        iLief% = Val(Left$(h$, 3))
        If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
            LetztGesendete$(j%) = h$
            If (j% > 0) Then Load mnuGesendetInd(j%)
            
            lif.GetRecord (iLief% + 1)
            h2$ = RTrim$(lif.kurz)
            
            iRufzeit% = Val(Mid$(h$, 4, 4))
            h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
            
            If (InStr(h$, "m.") > 0) Then h2$ = h2$ + "  manuell"
            
            mnuGesendetInd(j%).Caption = h2$
            j% = j% + 1
        End If
        
    End If
Next i%

Call DefErrPop
End Sub

Sub ErzeugeSchwellwerteMenu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeSchwellwerteMenu")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, iLief%, iRufzeit%
Dim h$, h2$

On Error Resume Next
For i% = 1 To 9
    Unload mnuSchwellwertInd(i%)
Next i%
On Error GoTo DefErr


j% = 0
For i% = 9 To 1 Step -1
    h$ = Dir("winw\*.sw" + Format(i%, "0"))
    If (h$ <> "") Then
        iLief% = Val(Left$(h$, 3))
        If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
            LetztSchwellwerte$(j%) = h$
            If (j% > 0) Then Load mnuSchwellwertInd(j%)
            
            lif.GetRecord (iLief% + 1)
            h2$ = RTrim$(lif.kurz)
            
            iRufzeit% = Val(Mid$(h$, 4, 4))
            h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
            
            If (InStr(h$, "m.") > 0) Then h2$ = h2$ + "  manuell"
            
            mnuSchwellwertInd(j%).Caption = h2$
            j% = j% + 1
        End If
        
    End If
Next i%

Call DefErrPop
End Sub

Sub ErzeugeDirektbezugMenu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeDirektbezugMenu")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, iLief%, iRufzeit%
Dim h$, h2$

On Error Resume Next
For i% = 1 To 9
    Unload mnuDirektbezugInd(i%)
Next i%
On Error GoTo DefErr


j% = 0
For i% = 9 To 1 Step -1
    h$ = Dir("winw\*.db" + Format(i%, "0"))
    If (h$ <> "") Then
        iLief% = Val(Left$(h$, 3))
        If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
            LetztDirektbezuege$(j%) = h$
            If (j% > 0) Then Load mnuDirektbezugInd(j%)
            
            lif.GetRecord (iLief% + 1)
            h2$ = RTrim$(lif.kurz)
            
            iRufzeit% = Val(Mid$(h$, 4, 4))
            h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
            
            If (InStr(h$, "m.") > 0) Then h2$ = h2$ + "  manuell"
            
            mnuDirektbezugInd(j%).Caption = h2$
            j% = j% + 1
        End If
        
    End If
Next i%

Call DefErrPop
End Sub

Sub ErzeugeRueckrufMenu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeRueckrufMenu")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, gef%, iLief%, RufLiefs%()
Dim h$

On Error Resume Next
For i% = 1 To 19
    Unload mnuRueckrufeInd(i%)
Next i%
On Error GoTo DefErr

With lstSortierung
    ReDim RufLiefs%(0)
    .Clear
    
    For i% = 0 To (AnzRufzeiten% - 1)
        gef% = False
        iLief% = Rufzeiten(i%).Lieferant
        For j% = 1 To UBound(RufLiefs%)
            If (RufLiefs%(j%) = iLief%) Then
                gef% = True
                Exit For
            End If
        Next j%
        If (gef% = False) Then
            k% = UBound(RufLiefs%) + 1
            ReDim Preserve RufLiefs%(k%)
            RufLiefs%(k%) = iLief%
            
            lif.GetRecord (iLief% + 1)
            h$ = RTrim$(lif.kurz)
            h$ = h$ + " (" + Mid$(Str$(iLief%), 2) + ")"
            .AddItem h$
        End If
    Next i%
    
    For i% = 0 To (.ListCount - 1)
        If (i% > 0) Then Load mnuRueckrufeInd(i%)
        .ListIndex = i%
        mnuRueckrufeInd(i%).Caption = .text
    Next i%
End With
    
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

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
DefErrFnc ("Form_LinkExecute")
DefErrMod (DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, kAsc%
Dim hWnd As Long, l As Long
Dim ScrAktForm$, n2 As String

Cancel = 0
            
ScrAktForm$ = ""
On Error Resume Next
ScrAktForm$ = Screen.ActiveForm.Name
On Error GoTo DefErr

hWnd = GetForegroundWindow()
n2 = Space(255)
l = GetWindowText(hWnd, n2, Len(n2))
If (Left$(n2, 16) = "Warenübernahme -") Then
'If (ScrAktForm$ = Me.Name) Then
    For i% = 1 To Len(CmdStr)
        kAsc% = Asc%(Mid$(CmdStr, i%, 1))
        If ((kAsc% >= 48) And (kAsc% <= 57)) Then EingabeStr$ = EingabeStr$ + Chr$(kAsc%)
    Next i%
    cmdOk(0).Value = True
Else
    Call ForwardDDE(CmdStr)
End If

Call DefErrPop
End Sub




