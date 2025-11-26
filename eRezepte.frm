VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm_eRezepte 
   Caption         =   "eRezepte"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   13815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "eRezepte.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   6690
   ScaleWidth      =   13815
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   11040
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2400
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
      Left            =   10920
      ScaleHeight     =   360
      ScaleWidth      =   1095
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstSortierung 
      Height          =   300
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picAnimationBack 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   2400
      ScaleHeight     =   2370
      ScaleWidth      =   5625
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   5655
      Begin ComCtl2.Animation aniAnimation 
         Height          =   1095
         Left            =   2280
         TabIndex        =   14
         Top             =   720
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
         TabIndex        =   15
         Top             =   180
         Width           =   5355
      End
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
      Left            =   3240
      ScaleHeight     =   360
      ScaleWidth      =   1335
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
      Begin VB.ComboBox cboAuswahlRezeptTyp 
         Height          =   360
         Left            =   6000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   33
         Top             =   120
         Width           =   1200
      End
      Begin VB.ComboBox cboAuswahlDatum 
         Height          =   360
         Left            =   4200
         Style           =   2  'Dropdown-Liste
         TabIndex        =   21
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtSuchen 
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
         Left            =   480
         TabIndex        =   32
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdSuchen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         Picture         =   "eRezepte.frx":030A
         Style           =   1  'Grafisch
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox chkAuswertung 
         Caption         =   "NUR &Rezepturen"
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboAuswertung 
         Height          =   360
         Index           =   0
         Left            =   480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtAuswertung 
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
         Index           =   0
         Left            =   2880
         TabIndex        =   27
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAuswertung 
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
         Index           =   1
         Left            =   4440
         TabIndex        =   26
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboAuswahl 
         Height          =   360
         Index           =   0
         Left            =   2760
         Style           =   2  'Dropdown-Liste
         TabIndex        =   20
         Top             =   120
         Width           =   1200
      End
      Begin SHDocVwCtl.WebBrowser wbAbgabedaten 
         Height          =   2535
         Left            =   5640
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   5175
         ExtentX         =   9128
         ExtentY         =   4471
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser wbVerordnung 
         Height          =   2535
         Left            =   1080
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   5175
         ExtentX         =   9128
         ExtentY         =   4471
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.PictureBox picBestellWerte 
         AutoRedraw      =   -1  'True
         Height          =   615
         Left            =   960
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
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
         Visible         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid flxSummen 
         Height          =   540
         Left            =   6840
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   953
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
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxQuittieren 
         Height          =   375
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
      End
      Begin VB.Label lblAuswahl 
         Caption         =   "&Anzeige"
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
         Left            =   0
         TabIndex        =   19
         Top             =   0
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
         Left            =   0
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
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   9615
      End
   End
   Begin VB.Label lblchkAuswertung 
      Caption         =   "AAAA"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   29
      Top             =   0
      Width           =   1095
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   4
      Left            =   11520
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   3
      Left            =   11520
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   2
      Left            =   11520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
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
         NumListImages   =   27
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":095F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":0BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":0F0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":119D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":142F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":16C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":19DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":1C6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":1EFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2011
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":22A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":25BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":26CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2961
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2A73
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2B85
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":2DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":303B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":32CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":355F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":37F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":3A83
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":3D9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":40B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":43D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":46EB
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
         NumListImages   =   27
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":4A05
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":4B17
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":4DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":503B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":514D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":525F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":54F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":5783
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":5895
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":5B27
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":5DB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":60D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":61E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":6477
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":6709
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":699B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":6C2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":6EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":7151
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":73E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":74F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":7607
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":7719
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":7A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":7D4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":8067
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "eRezepte.frx":8381
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDrucken 
         Caption         =   "&Drucken"
         Begin VB.Menu mnuDruckenInd 
            Caption         =   "&Aktuelle Ansicht"
            Index           =   0
         End
         Begin VB.Menu mnuDruckenInd 
            Caption         =   "&Zusammenfassung"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDruckenInd 
            Caption         =   "Alle Kunden &einzeln"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDruckenInd 
            Caption         =   "Aktuelle Ansicht (E&xcel)"
            Index           =   3
         End
      End
      Begin VB.Menu mnuBeenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuBearbeiten 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   ""
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
         Caption         =   "eRezept zurückgeben"
         Index           =   3
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Dispensieren und Einreichen eRezept"
         Index           =   4
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Aktualisieren"
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
         Caption         =   "Ansicht eRezept"
         Index           =   9
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Alle Zeilen auswählen"
         Index           =   10
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Auswahl rücksetzen"
         Index           =   11
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Dispensieren"
         Index           =   12
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Vorabprüfung"
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
      Begin VB.Menu mnuDummy11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBearbeitenLayout 
         Caption         =   "La&yout editieren"
      End
      Begin VB.Menu mnuDummy12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKomfortSignatur 
         Caption         =   "&Komfort-Signatur"
         Begin VB.Menu mnuKomfortSignaturInd 
            Caption         =   "&Aktivieren"
            Index           =   0
         End
         Begin VB.Menu mnuKomfortSignaturInd 
            Caption         =   "&Check"
            Index           =   1
         End
         Begin VB.Menu mnuKomfortSignaturInd 
            Caption         =   "&Deaktivieren"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDummy13 
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
      Begin VB.Menu mnuNlToolbar 
         Caption         =   "&Symbolleiste"
         Visible         =   0   'False
         Begin VB.Menu mnuNlToolbarInd 
            Caption         =   "&kleine Symbole"
            Index           =   0
         End
         Begin VB.Menu mnuNlToolbarInd 
            Caption         =   "&mittlere Symbole"
            Index           =   1
         End
         Begin VB.Menu mnuNlToolbarInd 
            Caption         =   "&grosse Symbole"
            Index           =   2
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
Attribute VB_Name = "frm_eRezepte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "ABCanalyse"
Const INFO_SECTION = "Infobereich ABCanalyse"

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

'Dim InRowColChange%
Dim HochfahrenAktiv%
Dim ProgrammModus%

Dim ArtikelStatistik%

'Dim FD_OP As TI_Back.Fachdienst_OP

Dim SollStatus%

Dim SortCol%
Dim SortModus%

Dim SortCols%(5)
Dim AnzSortCols%
Dim SortTypen$

Dim ShowSummen%

Private Const DefErrModul = "EREZEPTE.FRM"

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

KeinRowColChange% = True
Call opBereich.WechselModus(NeuerModus)
KeinRowColChange% = False
Call opToolbar.ShowToolbar
ProgrammModus% = NeuerModus%

'Select Case NeuerModus%
'    Case 0
'        mnuDatei.Enabled = True
'        mnuBearbeiten.Enabled = True
'        mnuAnsicht.Enabled = True
'
'        On Error Resume Next
'        For i% = MENU_F2 To MENU_SF9
'            mnuBearbeitenInd(i%).Enabled = EnabledArray%(i%)
'        Next i%
'        On Error GoTo DefErr
'
'        mnuBearbeitenLayout.Checked = False
'
'        cmdOk(0).Default = True
'        cmdEsc(0).Cancel = True
'
'        flxarbeit(0).BackColorSel = vbHighlight
'        flxInfo(0).BackColorSel = vbHighlight
'
'        h$ = Me.Caption
'        ind% = InStr(h$, " (EDITIER-MODUS)")
'        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
'        Me.Caption = h$
'    Case 1
'        mnuDatei.Enabled = False
'        mnuBearbeiten.Enabled = True
'        mnuAnsicht.Enabled = False
'
'        For i% = MENU_F2 To MENU_SF9
'            EnabledArray%(i%) = mnuBearbeitenInd(i%).Enabled
'        Next i%
'
'        mnuBearbeitenInd(MENU_F2).Enabled = True
'        mnuBearbeitenInd(MENU_F3).Enabled = False
'        mnuBearbeitenInd(MENU_F4).Enabled = False
'        mnuBearbeitenInd(MENU_F5).Enabled = True
'        mnuBearbeitenInd(MENU_F6).Enabled = False
'        mnuBearbeitenInd(MENU_F7).Enabled = False
'        mnuBearbeitenInd(MENU_F8).Enabled = True
'        mnuBearbeitenInd(MENU_F9).Enabled = False
'        mnuBearbeitenInd(MENU_SF2).Enabled = False
'        mnuBearbeitenInd(MENU_SF3).Enabled = False
'        mnuBearbeitenInd(MENU_SF4).Enabled = False
'        mnuBearbeitenInd(MENU_SF5).Enabled = False
'        mnuBearbeitenInd(MENU_SF6).Enabled = False
'        mnuBearbeitenInd(MENU_SF7).Enabled = False
'        mnuBearbeitenInd(MENU_SF8).Enabled = False
'
'        mnuBearbeitenLayout.Checked = True
'
'        cmdOk(0).Default = True
'        cmdEsc(0).Cancel = True
'
'        flxarbeit(0).BackColorSel = vbMagenta
'        With flxInfo(0)
'            .BackColorSel = vbMagenta
'            If (.Visible) Then
'                .SetFocus
'            End If
'        End With
'
'        h$ = Me.Caption
'        Me.Caption = h$ + " (EDITIER-MODUS)"
'End Select
'
'For i% = 0 To 7
'    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
'Next i%
'For i% = 8 To 15
'    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
'Next i%
'Call opToolbar.ShowToolbar
'
''Me.Caption = ProgrammName$ + lblArbeit(NeuerModus%).Caption
'
'ProgrammModus% = NeuerModus%

Call DefErrPop
End Sub

Private Sub chkAuswertung_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAuswertung_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With cboAuswertung(index)
    If (.Visible) Then
        Call AuswahlBefüllen
    End If
End With

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
Dim i%, row%, col%, erg%
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
'    frm_eRezeptEdit.Show 1
    With flxarbeit(0)
       frmAction.txtRezeptNr.text = "e" + .TextMatrix(.row, 2)
       
       eRezeptListe = ","
       For i = .FixedRows To .Rows - 1
            eRezeptListe = eRezeptListe + "E" + .TextMatrix(i, 2) + ","
       Next i
'       frmAction.txtRezeptNr.text = "e1r"
    End With
    Unload Me
ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
    With flxInfo(0)
        row% = .row
        col% = .col
        h$ = RTrim(.text)
        
        Dim sAktion$
        sAktion = .TextMatrix(row, 2)
        If (sAktion = "Vorab") Or (sAktion = "Ergebnis") Then
            XmlResponse = .TextMatrix(row, 5)
            frmFiveRxRueckmeldung.Show 1
        End If
    End With
    If (col% = 0) Then
'        Call ActProgram.cmdOkClick(h$)
    End If
ElseIf (ActiveControl.Name = txtSuchen.Name) Then
    With flxarbeit(0)
        Dim gef As Boolean
        gef = False
        For i = (.row + 1) To (.Rows - 1)
            If (InStr(LCase(.TextMatrix(i, SortCol)), LCase(txtSuchen.text)) > 0) Then
                .row = i
                .col = 0
                .ColSel = .Cols - 1
                gef = True
                .SetFocus
                Exit For
            End If
        Next i
        If (gef = False) Then
            For i = .FixedRows To .row
                If (InStr(LCase(.TextMatrix(i, SortCol)), LCase(txtSuchen.text)) > 0) Then
                    .row = i
                    .col = 0
                    .ColSel = .Cols - 1
                    .SetFocus
                    Exit For
                End If
            Next i
        End If
    
        If (.row < .TopRow) Then
            .TopRow = .row
        Else
            While (.TopRow + opBereich.ArbeitAnzZeilen - 1 <= .row)
                .TopRow = .row
            Wend
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub cboAuswahl_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAuswahl_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With cboAuswahl(index)
    If (.Visible) Then
        Dim ind%, ind2%
        ind = .ListIndex
        ind2 = 2
        If (ind = 1) Then
            ind2 = 0
        ElseIf (ind = 2) Or (ind = 3) Then
            ind2 = 1
        ElseIf (ind = 0) Or (ind = .ListCount - 1) Then
            ind2 = 3
        End If
        With cboAuswahlDatum
            .ListIndex = ind2
            .Enabled = True ' (ind2 > 0)
        End With
        
        Call AuswahlBefüllen
    End If
End With

Call DefErrPop
End Sub

Private Sub cboAuswahlDatum_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAuswahlDatum_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With cboAuswahlDatum
    If (.Visible) Then
        Dim ind%
        ind = .ListIndex
        
        cboAuswertung(0).ListIndex = IIf(ind <= 3, 14, 3)
        cboAuswertung(0).Enabled = (ind <= 3)
        txtAuswertung(1).Enabled = (ind <= 3)
    End If
End With

Call DefErrPop
End Sub

Private Sub cboAuswahlRezeptTyp_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAuswahlRezeptTyp_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With cboAuswahlRezeptTyp
    If (.Visible) Then
        Call AuswahlBefüllen
    End If
End With

Call DefErrPop
End Sub

Private Sub cmdSuchen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdSuchen_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

txtSuchen.SetFocus
cmdOk(0).Value = True

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

Private Sub flxarbeit_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(flxarbeit(index), picBack(index), Source, x, y)
Call DefErrPop
End Sub

Private Sub flxarbeit_DblClick(index As Integer)
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

Private Sub flxarbeit_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_Down")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Shift And vbShiftMask) <> 0 And (Shift And vbAltMask) <> 0 Then
    If KeyCode = vbKeyQ Then
        Call CreateQR
    End If
End If
If (ChefModus) And (KeyCode = vbKeyF9) Then
    ShowSummen = True
    flxSummen.Visible = True
End If

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

Private Sub flxInfo_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
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

Call opToolbar.Move(flxInfo(index), picBack(index), Source, x, y)

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

With flxInfo(index)
    .row = .FixedRows
    .col = 0
    .ColSel = .Cols - 1
End With

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

'On Error Resume Next
'If (index = 0) Then
'    Call FormKurzInfo
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
Static aRow%

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (index = 0) Then
    If ((flxarbeit(0).Redraw = True) And (KeinRowColChange% = False)) Then
        If (flxarbeit(0).row <> aRow) Then
            Call FormKurzInfo
            aRow = flxarbeit(0).row
    '        Call HighlightZeile
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
Dim i%, j%, ind%, wi%, WoTag%
Dim l&, KuNr&, AltKuNr&
Dim h$

HochfahrenAktiv% = True

eRezeptTaskId = ""
eRezeptListe = ""

If (FDok = 0) Then
    Set FD_OP = New TI_Back.Fachdienst_OP
    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), (wpara.EntwicklungsUmgebung = 0))) Then
        FDok = True
    End If
End If

If (para.Newline) Then
    WindowState = frmAction.WindowState
    Width = frmAction.Width
    Height = frmAction.Height
    Top = frmAction.Top
    Left = frmAction.Left
    Call wpara.ControlBorderless(Me, 3, wpara.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
    mnuDatei.Caption = Space(2) + mnuDatei.Caption
Else
    Width = frmAction.ScaleWidth - (180 * wpara.BildFaktor)
    Height = frmAction.ScaleHeight - (180 * wpara.BildFaktor)
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End If

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
    .Visible = True
End With

h$ = "2314241625151717,1418,1519,1320"
h$ = ",1418,1519,1323,0525"
 
'l& = WritePrivateProfileString("Allgemein", "Toolbar", Str$(2), CurDir + "\newline.ini")

Set opToolbar = New clsToolbar
'Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION, h$)
If (para.Newline) Then
    Call opToolbar.InitToolbar(Me, App.EXEName, INI_SECTION, h)
Else
    Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION, h)
End If

cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
cmdToolbar(1).ToolTipText = "F2"
cmdToolbar(2).ToolTipText = "F3"
cmdToolbar(3).ToolTipText = "F4"
cmdToolbar(4).ToolTipText = "F5"
cmdToolbar(5).ToolTipText = "F6"
cmdToolbar(6).ToolTipText = "F7"
cmdToolbar(7).ToolTipText = "F8"
cmdToolbar(8).ToolTipText = "F9"
cmdToolbar(9).ToolTipText = "shift+F2"
cmdToolbar(10).ToolTipText = "shift+F3"
cmdToolbar(11).ToolTipText = "shift+F4"
cmdToolbar(12).ToolTipText = "shift+F5"
cmdToolbar(13).ToolTipText = "shift+F6"
cmdToolbar(14).ToolTipText = "shift+F7"
cmdToolbar(15).ToolTipText = "shift+F8"
cmdToolbar(16).ToolTipText = "shift+F9"
'cmdToolbar(19).ToolTipText = "Programm beenden"

On Error Resume Next
For i% = MENU_F2 To MENU_SF9
    If (i% < MENU_SF2) Then
        j% = i% + 1
    Else
        j% = i%
    End If
    h$ = mnuBearbeitenInd(i%).Caption
    ind% = InStr(h$, "&")
    If (ind% > 0) Then
        h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
    End If
    cmdToolbar(j%).ToolTipText = cmdToolbar(j%).ToolTipText + " " + h$
Next i%
On Error GoTo DefErr

ArtikelStatistik = False

ShowSummen = Not (ChefModus)

Call wpara.InitFont(Me)
Call HoleIniWerte

Set InfoMain = New clsInfoBereich
Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, INFO_SECTION, 0, 5)
Call InfoMain.ZeigeInfoBereich("", False)
'Call ZeigeInfoBereichAdd(0)
flxInfo(0).row = 0
flxInfo(0).col = 0

Set opBereich = New clsOpBereiche
Call opBereich.InitBereich(Me, opToolbar)
opBereich.AutoRedraw = 0
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = True
opBereich.ArbeitWasDarunter = True
opBereich.InfoTitel = False
opBereich.InfoZusatz = ArtikelStatistik%
opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = 0 '-2

mnuZusatzInfo.Checked = ArtikelStatistik%

ProgrammModus% = 0

With flxarbeit(0)
    .Cols = 12
    .Rows = 0
    .Rows = 2
    .FixedRows = 1
    
    Dim sFormatStr$
    sFormatStr = "|<AbgabeDat|<TaskId|<RezeptNr|<PrescriptionId|^VO#|^ABG#|<Pzn|<Packung|||<Dar|^N|<ChargenNr|<Dosierung|^AI#|<KK|<KostentraegerName|"
    'wegen des irrtümlichen Anzeige der unsichtbaren Spalten rechts in PatientName
    'sFormatStr = sFormatStr + "<ZuzahlungsStatus|^ZuzFr|<Versicherter|<PatientName|<RezeptTyp|<VerordnungsTyp|<Kategorie|<AccessCode|<Secret|"
    'sFormatStr = sFormatStr + "<Dosieranweisung|^AutIdem|PatientTyp|"
    sFormatStr = sFormatStr + "<ZuzahlungsStatus|^ZuzFr|<Versicherter|<PatientName||||||"
    sFormatStr = sFormatStr + "|||"
    sFormatStr = sFormatStr + "^Impf|^BVG|^Noct|<AbrechBis|<EinreichBis|<BundleKBV|Quittung|<Kz|<AbrufDat|<VerkaufDat|<DispenseDat|<EinreichDat|>GesBrutto|>Zuzahl|>ALBVVG|<Status||"
    .FormatString = sFormatStr

    SortTypen = "3201011100000003100300000000003332200022221110" + String(20, "0")

    .GridLines = flexGridFlat
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightWithFocus
End With

With flxInfo(0)
    .Cols = 6
    .Rows = 0
    .Rows = 2
    .FixedRows = 1
    
'    sFormatStr = "|<Gescannt|<TaskId|<PrescriptionId|<AnzahlVerordnetePackungen|<Pzn|<Packung|Packungsgroesse|<Einheit|<Darreichungsform|^N123|^AI#|<KK|<KostentraegerName|"
    sFormatStr = "^|<Datum|<Aktion|<Ergebnis||"
    .FormatString = sFormatStr

    .GridLines = flexGridFlat
    .SelectionMode = flexSelectionByRow
    .HighLight = flexHighlightWithFocus
    
    .ScrollBars = flexScrollBarVertical
End With

'With flxarbeit(0)
'    .Cols = 18
'    .Rows = 2
'    .FixedRows = 1
'    .FormatString = "|^" + Chr$(214) + "|>LfdNr|<Beleg|^|<Datum|<Zeit|>KNr|<Kurz|>Zeilen|>Rezepte|>Wert|>Wert inkl.|^SammelK|>|>|>|"
'
'    .GridLines = flexGridFlat
'    .SelectionMode = flexSelectionByRow
'    .HighLight = flexHighlightWithFocus
'End With

'With flxInfo(0)
'    .Cols = 21
'    .Rows = 0
'    .Rows = 2
'    .FixedRows = 1
'    .FormatString = "<PZN|^Anzahl|<Name|>Menge|^Meh|>Preis|>Preis inkl.|>Rabatt (%)|>Wert|>Wert inkl.|>MwSt|>WG|<Typ|||||||"
'End With

AltKuNr = 0
With cboAuswahl(0)
    .Clear
    .AddItem "" + Space(100) + "0"
    .AddItem "Eingescannt" + Space(100) + "2"
    .AddItem "Dispensiert" + Space(100) + "3"
    .AddItem "Dispensiert (HBA Signierung)" + Space(100) + "3"
    .AddItem "Eingereicht" + Space(100) + "4"
    .AddItem "Fehler" + Space(100) + "5"
    .AddItem "Abrechenbar" + Space(100) + "6"
    .AddItem "Ohne weitere Bearbeitung" + Space(100) + "7"
    .AddItem "Unbearbeitet" + Space(100) + "-1"
'    .AddItem "Globale Suche" + Space(100) + "0"
    
    .ListIndex = 1
End With
With cboAuswahlDatum
    .Clear
    .AddItem "Abgabe-Datum"
    .AddItem "Dispensier-Datum"
    .AddItem "Einreich-Datum"
    .AddItem "Abruf-Datum"
    
    .AddItem "ChargenNr"
    .AddItem "KK"
    .AddItem "KostentraegerName"
    .AddItem "Packung"
    .AddItem "PatientName"
    .AddItem "Pzn"
    .AddItem "RezeptNr"
    .AddItem "TaskId"
    .AddItem "Versicherter"
    
    .ListIndex = 0
    .Enabled = False
End With
With cboAuswertung(0)
    .Clear
    .AddItem ""
    
    .AddItem "<"
    .AddItem "<="
    .AddItem "="
    .AddItem "<>"
    .AddItem ">="
    .AddItem ">"
    .AddItem String(50, "-")
    .AddItem "Zwischen"
    .AddItem String(50, "-")
    .AddItem "Heute"
    .AddItem "Gestern"
    .AddItem "diese Woche"
    .AddItem "letzte Woche"
    .AddItem "letzte 7 Tage"
    .AddItem "dieses Monat"
    .AddItem "letztes Monat"
    .AddItem "letzte 4 Wochen"
    .AddItem "letzte 30 Tage"
    .AddItem "dieses Quartal"
    .AddItem "letztes Quartal"
    .AddItem "dieses Jahr"
    .AddItem "letztes Jahr"
    .AddItem "letzte 12 Monate"
    .AddItem "letzte 52 Wochen"
    .AddItem "letzte 365 Tage"
    
'    .ListIndex = 14
    
'    txtAuswertung(1).Visible = True
'    WoTag = Weekday(Now, vbMonday)
'    txtAuswertung(0).text = Format(Now - 6, "DDMMYY")
'    txtAuswertung(1).text = Format(Now, "DDMMYY")

End With
With cboAuswahlRezeptTyp
    .Clear
    .AddItem "Alle Rezepte"
    .AddItem "Rezepturen"
    .AddItem "Pharm.DL"
    
    .ListIndex = 0
'    .Enabled = False
End With
'Call AuswahlBefüllen
        
mnuBearbeitenInd(MENU_F3).Enabled = True
mnuBearbeitenInd(MENU_F4).Enabled = False
mnuBearbeitenInd(MENU_F7).Enabled = True
mnuBearbeitenInd(MENU_SF3).Enabled = True  'ParamSammelbeleg%
mnuBearbeitenInd(MENU_SF4).Enabled = True  ' ParamSammelbeleg%
mnuBearbeitenInd(MENU_SF6).Enabled = True
mnuBearbeitenInd(MENU_SF7).Enabled = False
mnuBearbeitenInd(MENU_SF8).Enabled = False
For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%

mnuBearbeitenLayout.Enabled = True


'mnuBearbeitenZusatzInd(0).Caption = "Alte Eintragungen aus Merkzettel entfernen"
'mnuBearbeitenZusatzInd(0).Enabled = True
'Load mnuBearbeitenZusatzInd(1)
'mnuBearbeitenZusatzInd(1).Caption = "Lagernde Artikel aus Merkzettel entfernen"
'mnuBearbeitenZusatzInd(1).Enabled = True

If (para.Newline) Then
    mnuNlToolbar.Visible = True
    mnuNlToolbarInd(opToolbar.Size).Checked = True
    mnuToolbar.Visible = False
End If


Call WechselModus(1)
Call WechselModus(0)

Call InitAnimation

'HochfahrenAktiv% = False
picBack(0).Visible = True

Call WheelHook(Me.hWnd)
  
Call DefErrPop
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call WheelUnHook(Me.hWnd)

'Set FD_OP = Nothing

End Sub

Private Sub mnuKomfortSignaturInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuKomfortSignaturInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim h, IccSN As String

IccSN = FD_OP.GetHBA_IccSN
If (IccSN = "") Then
    Call MessageBox("Achtung: Probleme mit der HBA-Karte! Funktion nicht durchführbar!", vbCritical)
    Call DefErrPop: Exit Sub
End If

h = "Komfort-Signatur für die gesteckte HBA-Karte (" + IccSN + ") "
If (index = 0) Then
    If (MessageBox(h + " AKTIVIEREN?", vbYesNo Or vbDefaultButton1) = vbYes) Then
        Call FD_OP.Komfortsignatur_Aktivieren(Me.hWnd, IccSN)
'        MsgBox ("ok")
    End If
ElseIf (index = 1) Then
    If (MessageBox(h + " CHECKEN?", vbYesNo Or vbDefaultButton1) = vbYes) Then
        Call FD_OP.Komfortsignatur_Check(Me.hWnd, IccSN)
'        MsgBox ("ok")
    End If
ElseIf (index = 2) Then
    If (MessageBox(h + " DEAKTIVIEREN?", vbYesNo Or vbDefaultButton1) = vbYes) Then
        Call FD_OP.Komfortsignatur_Deaktivieren(Me.hWnd, IccSN)
'        MsgBox ("ok")
    End If
End If

Call DefErrPop
End Sub

Private Sub picBack_Paint(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picBack_Paint")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim x1&, x2&, y1&, Y2&, iAdd&, iAdd2&
Dim c As Control

Call wpara.picBackPaint(picBack(0), opBereich.InfoZusatz)
    
On Error Resume Next
For Each c In Controls
    If (c.tag <> "0") Then
        If (c.Container.Name = picBack(0).Name) Then
            If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                c.BackColor = vbWhite
                With c.Container
                    .ForeColor = RGB(180, 180, 180) ' vbWhite
                    .FillStyle = vbSolid
                    .FillColor = vbWhite

                    RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                End With
            End If
        End If
    End If
Next

Call DefErrPop
End Sub

Private Sub lblarbeit_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(lblArbeit(index), picBack(index), Source, x, y)
Call DefErrPop
End Sub

Private Sub lblInfo_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(lblInfo(index), picBack(index), Source, x, y)
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, row%, col%, ind%, erg%
Dim KuNr&, AltKuNr&
Dim h$, h2$, sTaskId$
Dim bErgebnis As Boolean
Dim eRezept2 As TI_Back.eRezept_OP

If (para.Newline) Then
    ind = index
    If (ind <= MENU_F9) Then
        ind = ind + 1
    End If
    Call opToolbar.Click(-ind)
End If

With flxarbeit(0)
'    pzn$ = .TextMatrix(.row, 0)
'    txt$ = Trim(.TextMatrix(.row, 2)) + "  " + Trim(.TextMatrix(.row, 3)) + .TextMatrix(.row, 4)
End With

Select Case index

    Case MENU_F2
'        If (ProgrammModus% = 1) Then
'            If (ActiveControl.Name = flxInfo(0).Name) Then
'                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
'                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
'                Call opBereich.RefreshBereich
'                Call FormKurzInfo
'            End If
'        Else
'
'
'            ParamSammelbeleg% = 3
'            ParamKunden$ = ""
'            ParamBelegTyp$ = "GR"
'
'            frmHolenParam.Show 1
'
'            If (ParamBelegTyp$ <> "") Then
'
'                If (ParamKundenSort) Then
'                    With cboAuswahl(0)
'                        .Clear
'                        .AddItem "Zusammenfassung"
'                        .AddItem String(50, "-")
'                    End With
'
'                    Call AuswahlBefüllen
'                Else
'                    Load frmBelegHolen
'                    Unload frmBelegHolen
'
'                    If (frmAction.lstBelegIds.ListCount > 0) Then
'                        AltKuNr = 0
'                        With cboAuswahl(0)
'                            .Clear
'                            .AddItem "Zusammenfassung"
'                            .AddItem String(50, "-")
'                            For i% = 1 To frmAction.lstBelegIds.ListCount
'                                KuNr = Val(Left$(frmAction.lstBelegIds.List(i% - 1), 6))
'                                If (KuNr <> AltKuNr) Then
'                                    SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(KuNr)
'                                    #If (KUNDEN_SQL = -1) Then
'                                        On Error Resume Next
'                                        KundenRec.Close
'                                        Err.Clear
'                                        On Error GoTo DefErr
'                                        KundenRec.open SQLStr, KundenAdoDB.ActiveConn
'                                    #Else
'                                        Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
'                                    #End If
'                                    h$ = ""
'                                    If (KundenRec.EOF = False) Then
'                                        h$ = Trim(CheckNullStr(KundenRec!VorName))
'                                        h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Name)))
'                                    End If
'                                    cboAuswahl(0).AddItem h$ + " (" + Format$(KuNr, "0") + ")"
'                                    AltKuNr = KuNr
'                                End If
'                            Next i%
'                            .ListIndex = 0
'                        End With
'
'                        Call AuswahlBefüllen
'                    End If
'
'                End If
'            End If
'        End If
    
    Case MENU_F3
        Call AuswahlSortieren
    
    Case MENU_F5
'        If (ProgrammModus% = 1) Then
'            If (ActiveControl.Name = flxInfo(0).Name) Then
'                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
'                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
'                Call opBereich.RefreshBereich
'                Call FormKurzInfo
'            End If
'        End If
        If (SollStatus = 2) Then
            If (MessageBox("Aktuelles eRezept zurückgeben?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
                With flxarbeit(0)
                    sTaskId = .TextMatrix(.row, 2)
                End With
                If (FDok = 0) Then
                    Set FD_OP = New TI_Back.Fachdienst_OP
                    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
                        FDok = True
                    End If
                End If
                
                Dim eRezept As New TI_Back.eRezept_OP
                Call eRezept.New2(sTaskId)
                MsgBox ("Zurückgeben: " + CStr(FD_OP.TI_RezeptZurueckgeben(eRezept.TaskId, eRezept.Secret)))
                Call AuswahlBefüllen
            End If
        ElseIf (SollStatus = 4) Or (SollStatus = 5) Or (SollStatus = 6) Then
            If (MessageBox("Aktuelles eRezept beim ARZ STORNIEREN?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
                With flxarbeit(0)
                    sTaskId = .TextMatrix(.row, 2)
                End With
    
                erg = HoleErgebnisse(sTaskId, True)
                If (erg = 5) Then   'STORNIERT
                    SQLStr = "UPDATE TI_eRezepte SET OpStatus=" + "3"
                    SQLStr = SQLStr + " WHERE TaskId='" + sTaskId + "'"
                    VerkaufAdoDB.ActiveConn.Execute (SQLStr)
                    Call MessageBox("eRezept erfolgreich beim ARZ STORNIERT !", vbInformation)
                    
                    Call AuswahlBefüllen
                End If
            End If
        ElseIf (SollStatus = 3) Then
            With flxarbeit(0)
                .TextMatrix(.row, 0) = Chr(214)
                
                Dim iAnzUmgespeichert%
                iAnzUmgespeichert = 0
                For i% = .FixedRows To (.Rows - 1)
                    If (.TextMatrix(i%, 0) <> "") Then
                        sTaskId = .TextMatrix(i, 2)
            
                        SQLStr = "UPDATE TI_eRezepte SET OpStatus=" + "7"
                        SQLStr = SQLStr + " WHERE TaskId='" + sTaskId + "'"
                        'MsgBox (SQLStr)
                        VerkaufAdoDB.ActiveConn.Execute (SQLStr)
                        
                        iAnzUmgespeichert = iAnzUmgespeichert + 1
                    End If
                Next i%
            End With
            
            h = "Anzahl umgespeicherter eRezepte: " + CStr(iAnzUmgespeichert) + vbCrLf + vbCrLf
            Call MessageBox(h, vbInformation)
            
            Call AuswahlBefüllen
        End If

'    Case MENU_F6
''        MySendKeys "%DD", True
'        With flxarbeit(0)
'            sTaskId = .TextMatrix(.row, 2)
'        End With
'        If (FDok = 0) Then
'            Set FD_OP = New TI_Back.Fachdienst_OP
'            If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
'                FDok = True
'            End If
'        End If
'
'        Set eRezept2 = New TI_Back.eRezept_OP
'        bErgebnis = True
'        Call eRezept2.New2(sTaskId)
'        If (SollStatus = 4) Then
'            MsgBox ("Ergebnis: " + CStr(HoleErgebnisse(sTaskId)))
'        Else
'            If (SollStatus < 3) Then
'                If (FD_OP.TI_RezeptAbgeben(eRezept2.TaskId, eRezept2.Secret)) Then
'                    Call MessageBox("eRezept erfolgreich DISPENSIERT !", vbInformation)
'                Else
'                    Call MessageBox("Probleme beim Dispensieren !", vbCritical)
'                    bErgebnis = False
'                End If
'            End If
'            If (bErgebnis) Then
'                Set eRezept2 = New TI_Back.eRezept_OP
'                Call eRezept2.New2(sTaskId)
'                If (FD_OP.TI_RezeptAbrechnen(eRezept2, False)) Then
'    '                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
'                    bErgebnis = VorabPruefung(sTaskId, False)
'                Else
'                    Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
'                End If
'            End If
'        End If
'        Call AuswahlBefüllen
    Case MENU_F6
        If (FDok = 0) Then
            Set FD_OP = New TI_Back.Fachdienst_OP
            If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
                FDok = True
            End If
        End If
        
        With flxarbeit(0)
            .TextMatrix(.row, 0) = Chr(214)
            
            Dim iAnzGesamt%, iAnzDispensiert%, iAnzHBA%, iAnzEingereicht%, iAnzStatus%(8)
            iAnzGesamt = 0
            For i% = .FixedRows To (.Rows - 1)
                If (.TextMatrix(i%, 0) <> "") Then
                    iAnzGesamt = iAnzGesamt + 1
                End If
            Next i
            bEinzelErgebnis = (iAnzGesamt <= 1)
            
            
            iAnzDispensiert = 0
            iAnzHBA = 0
            iAnzEingereicht = 0
            For i = 0 To 8
                iAnzStatus(i) = 0
            Next i
            
            
            Set eRezept2 = New TI_Back.eRezept_OP
            If (SollStatus = 6) Or (SollStatus = 7) Then
                lstSortierung.Clear
                For i% = .FixedRows To (.Rows - 1)
                    If (.TextMatrix(i%, 0) <> "") Then
                        sTaskId = .TextMatrix(i, 2)
            
                        bErgebnis = True
                        Call eRezept2.New2(sTaskId)
                        
                        h = ""
                        With eRezept2.Patient.Name
                            h2 = .Nachname_ohne_Vor_und_Zusatz + ", " + .Vorname
                        End With
                        h2 = Left(h2 + Space(100), 100)
                        h = h + h2
                        
                        With eRezept2.Arzt.Name
                            h2 = .Nachname_ohne_Vor_und_Zusatz + ", " + .Vorname
                        End With
                        h2 = Left(h2 + Space(100), 100)
                        h = h + h2
                        
                        h2 = Format(eRezept2.AbgabeDatum, "YYYYMMDD")
                        h = h + h2
                        
                        h2 = Format(eRezept2.Erstellungsdatum, "YYYYMMDD")
                        h = h + h2
                        
                        h = h + "@" + sTaskId
                        
                        lstSortierung.AddItem (h)
                    End If
                Next i
                With lstSortierung
                    If (.ListCount > 0) Then
                        Dim sLastItem As String
                        h = ""
                        h2 = .List(0)
                        ind = InStr(h2, "@")
                        If (ind > 0) Then
                            h2 = Left(h2, ind - 1)
                        End If
                        sLastItem = h2
                        
                        For i = 0 To (.ListCount - 1)
                            h2 = .List(i)
                            ind = InStr(h2, "@")
                            If (ind > 0) Then
                                sTaskId = Mid(h2, ind + 1)
                                h2 = Left(h2, ind - 1)
                            End If
                            If (h2 <> sLastItem) Then
                                Call FD_OP.PKV_Ausdruck(h)
                                h = ""
                            End If
                            h = h + sTaskId + ";"
                            sLastItem = h2
                        Next i
                        If (h <> "") Then
                            Call FD_OP.PKV_Ausdruck(h)
                        End If
                    End If
                End With
'                        Call FD_OP.PKV_Ausdruck(eRezept2, True)
            Else
                For i% = .FixedRows To (.Rows - 1)
                    If (.TextMatrix(i%, 0) <> "") Then
                        sTaskId = .TextMatrix(i, 2)
            
                        bErgebnis = True
                        Call eRezept2.New2(sTaskId)
                        If (SollStatus = 4) Then
    '                        MsgBox ("Ergebnis: " + CStr(HoleErgebnisse(sTaskId)))
                            erg = HoleErgebnisse(sTaskId)
                            iAnzStatus(erg) = iAnzStatus(erg) + 1
                        Else
                            If (SollStatus < 3) Then
                                If (FD_OP.TI_RezeptAbgeben(eRezept2.TaskId, eRezept2.Secret)) Then
                                    If (bEinzelErgebnis) Then
                                        If (eRezept2.Kostentraeger.Typ = "SEL") Or (eRezept2.RezeptTypId = 200) Then
                                            SQLStr = "UPDATE TI_eRezepte SET OpStatus=" + "7"
                                            SQLStr = SQLStr + " WHERE TaskId='" + sTaskId + "'"
                                            'MsgBox (SQLStr)
                                            VerkaufAdoDB.ActiveConn.Execute (SQLStr)
                                        End If
                                        Call MessageBox("eRezept erfolgreich DISPENSIERT !", vbInformation)
                                    Else
                                        iAnzDispensiert = iAnzDispensiert + 1
                                    End If
                                Else
                                    If (bEinzelErgebnis) Then
                                        Call MessageBox("Probleme beim Dispensieren !", vbCritical)
                                    End If
                                    bErgebnis = False
                                End If
                            End If
                            If (bErgebnis) Then
                                Set eRezept2 = New TI_Back.eRezept_OP
                                Call eRezept2.New2(sTaskId)
                                If (SollStatus < 3) And (eRezept2.QES) Then
                                    iAnzHBA = iAnzHBA + 1
                                Else
    '                                h = InputBox("Anzahl verordnete Packungen: ", "eRezept", eRezept2.Anzahl)
    '                                If (h <> "") Then
    '                                    eRezept2.Anzahl = h
    '                                End If
                                    If (FD_OP.TI_RezeptAbrechnen(eRezept2, False)) Then
                        '                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
                                        bErgebnis = VorabPruefung(sTaskId, False)
                                        If (bErgebnis) Then
                                            iAnzEingereicht = iAnzEingereicht + 1
                                        End If
                                    Else
                                        If (bEinzelErgebnis) Then
                                            Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i%
            End If
        End With
        
        If (bEinzelErgebnis) Then
        Else
            h = "Anzahl ausgewählte eRezepte: " + CStr(iAnzGesamt) + vbCrLf + vbCrLf
            If (SollStatus = 4) Then
                For i = 1 To 8
                    If (iAnzStatus(i) > 0) Then
                        h = h + FiveRxRezeptStatus(i - 1) + ": " + CStr(iAnzStatus(i)) + vbCrLf
                    End If
                Next i
            Else
                If (iAnzDispensiert > 0) Then
                    h = h + "Erfolgreich dispensiert: " + CStr(iAnzDispensiert) + vbCrLf
                End If
                If (iAnzEingereicht > 0) Then
                    h = h + "Eingereicht, LieferID erhalten: " + CStr(iAnzEingereicht) + vbCrLf
                End If
                If (iAnzHBA > 0) Then
                    h = h + vbCrLf + "Noch zu HBA-Signieren: " + CStr(iAnzHBA) + vbCrLf
                End If
            End If
            Call MessageBox(h, vbInformation)
        End If
        
        Call AuswahlBefüllen
    
    Case MENU_F7
        Call AuswahlBefüllen
        
    Case MENU_F8
'        If (ProgrammModus% = 1) Then
'            If (ActiveControl.Name = flxInfo(0).Name) Then
'                col% = flxInfo(0).col
'                If (col% > 0) And (col% Mod 2) Then
'                    row% = flxInfo(0).row
'                    If (InfoMain.Bezeichnung(row%, (col% - 1) \ 2) <> "") Then
'                        Call EditInfoName
'                    End If
'                End If
'            End If
'        End If
    
    Case MENU_SF2
'        wbVerordnung.Visible = Not (wbVerordnung.Visible)
'        wbAbgabedaten.Visible = Not (wbAbgabedaten.Visible)
'        Call FormKurzInfo
                
        With flxarbeit(0)
            eRezeptTaskId = .TextMatrix(.row, 2)
        End With
        If (eRezeptTaskId <> "") Then
            If (FDok = 0) Then
                Set FD_OP = New TI_Back.Fachdienst_OP
                If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
                    FDok = True
                End If
            End If
        
            frm_eRezeptEdit.Show 1
        End If
        
    Case MENU_SF3
        Call AuswahlSetzen
        
    Case MENU_SF4
        Call AuswahlSetzen(False)
    
    Case MENU_SF5, MENU_SF6
        With flxarbeit(0)
            sTaskId = .TextMatrix(.row, 2)
        End With
        If (FDok = 0) Then
            Set FD_OP = New TI_Back.Fachdienst_OP
            If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
                FDok = True
            End If
        End If
        
        bEinzelErgebnis = True
        Set eRezept2 = New TI_Back.eRezept_OP
        Call eRezept2.New2(sTaskId)
        If (index = MENU_SF5) Then
            If (FD_OP.TI_RezeptAbgeben(eRezept2.TaskId, eRezept2.Secret)) Then
                Call MessageBox("eRezept erfolgreich DISPENSIERT !", vbInformation)
            Else
                Call MessageBox("Probleme beim Dispensieren !", vbCritical)
            End If
        Else
'            If (SollStatus = 2) Or (SollStatus = 3) Or (SollStatus = 5) Then
'                If (FD_OP.TI_RezeptAbrechnen(eRezept2, (SollStatus = 2))) Then
''                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
'                    bErgebnis = VorabPruefung(sTaskId, (SollStatus = 2))
'                Else
'                    Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
'                End If
'            ElseIf (SollStatus = 4) Then
'                MsgBox ("Ergebnis: " + CStr(HoleErgebnisse(sTaskId)))
'            End If
            If (FD_OP.TI_RezeptAbrechnen(eRezept2, True)) Then
'                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
                bErgebnis = VorabPruefung(sTaskId, True)
            Else
                Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
            End If
        End If
        Call AuswahlBefüllen

    Case MENU_SF7
    
    Case MENU_SF8

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

'Private Sub mnuZusatzinfo_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("mnuZusatzInfo_Click")
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
'Dim h$
'
'If (mnuZusatzInfo.Checked) Then
'    mnuZusatzInfo.Checked = False
'Else
'    mnuZusatzInfo.Checked = True
'End If
'
'opBereich.InfoZusatz = mnuZusatzInfo.Checked
'
'If (opBereich.InfoZusatz) Then
'    h$ = "J"
'Else
'    h$ = "N"
'End If
'l& = WritePrivateProfileString(INI_SECTION, "ArtikelStatistik", h$, INI_DATEI)
'
'opBereich.RefreshBereich
'
'Call DefErrPop
'End Sub

Private Sub mnuBearbeitenZusatzInd_Click(index As Integer)
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

'LoeschModus% = index
'Call LoescheMerkzettel

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

If ((Shift And vbShiftMask) = vbShiftMask) And ((Shift And vbAltMask) = vbAltMask) And (KeyCode = 191) Then
    h = Trim(InputBox("OP-Status des Rezeptes: ", "Status des eRezeptes", CStr(SollStatus)))
    If (Len(h) = 1) Then
        If (InStr("1234567", h) > 0) Then
            If (Val(h) <> SollStatus) Then
                Dim sTaskId$
                With flxarbeit(0)
                    sTaskId = .TextMatrix(.row, 2)
                End With
    
                SQLStr = "UPDATE TI_eRezepte SET OpStatus=" + h
                SQLStr = SQLStr + " WHERE TaskId='" + sTaskId + "'"
                'MsgBox (SQLStr)
                VerkaufAdoDB.ActiveConn.Execute (SQLStr)
                Call AuswahlBefüllen
            End If
        End If
    End If
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

On Error Resume Next

If (HochfahrenAktiv%) And (Me.Visible = False) Then
    Call DefErrPop: Exit Sub
End If

If (Me.WindowState = vbMinimized) Then Call DefErrPop: Exit Sub

Call opBereich.ResizeWindow
If (HochfahrenAktiv%) Then
    HochfahrenAktiv% = 0
    DoEvents
    cboAuswertung(0).ListIndex = 14
    Call AuswahlBefüllen
End If

picSave.Visible = False

Call DefErrPop
End Sub

Private Sub mnuDruckenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDruckenInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

'With cboAuswahl(0)
'    If (index = 0) Then
'        Call Drucke_eRezepte
'    ElseIf (index = 1) Then
'        .ListIndex = 0
'        DoEvents
' '       Call DruckeABCanalyse
'    ElseIf (index = 2) Then
'        For i% = 2 To (.ListCount - 1)
'            .ListIndex = i%
'            DoEvents
'  '          Call DruckeABCanalyse
'        Next i%
'    Else
'        Call Export_eRezepte
'    End If
'End With

If (index = 0) Then
    Call Drucke_eRezepte
Else
    Call Export_eRezepte
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

Private Sub mnuZusatzInfo_Click()
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

Private Sub picBack_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
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
Call opToolbar.Move(picBack(index), picBack(index), Source, x, y)
Call DefErrPop
End Sub

'Private Sub picBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("picBack_MouseMove")
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
'If (picToolTip.Visible = True) Then
'    picToolTip.Visible = False
'End If
'
'Call DefErrPop
'End Sub

'Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("picToolbar_MouseDown")
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
'picToolbar.Drag (vbBeginDrag)
'opToolbar.DragX = x
'opToolbar.DragY = y
'Call DefErrPop
'End Sub

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
Dim row%, iRow%, iCol%, XML_ID%
Dim l&
Dim pzn$, sDatei$, sBundleKBV$, sTaskId$, sBundleDAV$
    
If (Me.Visible = False) Then Call DefErrPop: Exit Sub

If (flxarbeit(0).Rows < 2) Then Call DefErrPop: Exit Sub
row% = flxarbeit(0).row

If (wbVerordnung.Visible) Then
    sBundleKBV = flxarbeit(0).TextMatrix(row%, 35)
        
    sDatei = CurDir() + "\eVerordnung.xml"
    'XML_ID% = FileOpen(sDatei, "O")
    'Print #XML_ID%, sBundleKBV;
    'Close #XML_ID
    l = writeOut(sBundleKBV, sDatei)
    
    l = Transformieren(sDatei, CurDir() + "\ERP_Stylesheet.xslt", CurDir() + "\eVerordnung.htm")
    With wbVerordnung
        .Visible = False
    '    .Navigate ("about:blank")
        .Width = frmAction.Width '/ 2
        .Height = frmAction.Height '/ 2
        .Width = 950 * 15
        .Height = 550 * 15
                    
        .Top = flxarbeit(0).Top + flxarbeit(0).RowPos(flxarbeit(0).row) + flxarbeit(0).RowHeight(0)
        If (.Top + .Height > picBack(0).Height) Then
            .Top = 0
        End If
        .Left = 60  'flxarbeit(0).Left + flxarbeit(0).ColPos(2)
        
        .Navigate (CurDir() + "\eVerordnung.htm")
        .Visible = True
    End With
    
    Dim sVersionAbgabedaten As String
    Dim eRezept As New TI_Back.eRezept_OP
    sTaskId = flxarbeit(0).TextMatrix(row%, 2)
    Call eRezept.New2(sTaskId)
    sBundleDAV = FD_OP.TI_RezeptAbgabedaten(eRezept, sVersionAbgabedaten)
'    MsgBox (sVersionAbgabedaten)
  
    sDatei = CurDir() + "\eAbgabe.xml"
    l = writeOut(sBundleDAV, sDatei)
      
    l = Transformieren(sDatei, CurDir() + "\eAbgabedaten.xslt", CurDir() + "\eAbgabe.htm")
    With wbAbgabedaten
        .Visible = False
        .Width = 900 * 15
        .Height = 550 * 15
        .Left = wbVerordnung.Left + wbVerordnung.Width + 60
        .Top = wbVerordnung.Top
        .Navigate (CurDir() + "\eAbgabe.htm")
        .Visible = True
    End With
    
    With flxarbeit(0)
        .row = row
        .col = 0
        .ColSel = .Cols - 1
        .SetFocus
    End With
End If

'MsgBox ("ok")
Dim bOk%, iAktion%, ind%
Dim h$, sAktion$, sTag$
flxInfo(0).Rows = flxInfo(0).FixedRows
SQLStr = "SELECT * FROM TI_Aktionen"
SQLStr = SQLStr + " WHERE (TaskId='" + flxarbeit(0).TextMatrix(row, 2) + "')"
SQLStr = SQLStr + " ORDER BY AnlageDatum"
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'If (FabsErrf <> 0) Then
'    Call iMsgBox("keine passenden Rezepte gespeichert !")
'    Call DefErrPop: Exit Sub
'End If
Do
    If (VerkaufAdoRec.EOF) Then
        Exit Do
    End If
    
    sAktion = CheckNullStr(VerkaufAdoRec!Aktion)
    
    bOk = True
    If (UCase(sAktion) = "ABGEBEN") Then
        bOk = (CheckNullStr(VerkaufAdoRec!DispenseXML) <> "") And (CheckNullStr(VerkaufAdoRec!ResultXml) <> "")
    ElseIf (UCase(sAktion) = "ABRUF") Then
        bOk = (CheckNullStr(VerkaufAdoRec!BundleXml) <> "")
    End If
    
    h = IIf(bOk, Chr(214), "")
    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "DD.MM.YY HH:mm")
    h = h + vbTab + sAktion
    h = h + vbTab
    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "YYMMDDHHmmss") + Format(VerkaufAdoRec!Id, "000000")
    flxInfo(0).AddItem h, 1
    
    VerkaufAdoRec.MoveNext
Loop
VerkaufAdoRec.Close

SQLStr = "SELECT * FROM TI_FiveRx" '
SQLStr = SQLStr + " WHERE (TaskId='" + flxarbeit(0).TextMatrix(row, 2) + "')"
SQLStr = SQLStr + " ORDER BY AnlageDatum"
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
Do
    If (VerkaufAdoRec.EOF) Then
        Exit Do
    End If
    
    iAktion = CheckNullInt(VerkaufAdoRec!AktionInt)
    'MsgBox ("drin:" + CStr(iAktion))
    If (iAktion = 0) Or (iAktion = 2) Or (iAktion = 4) Then
        If (iAktion = 0) Then
            sAktion = "Vorab"
        ElseIf (iAktion = 2) Then
            sAktion = "Einreichung"
        Else
            sAktion = "Ergebnis"
        End If
    
        h = ""
        h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "DD.MM.YY HH:mm")
        h = h + vbTab + sAktion
        h = h + vbTab
        h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "YYMMDDHHmmss") + Format(VerkaufAdoRec!Id, "000000")
        flxInfo(0).AddItem h, 1
    Else
        sAktion = CheckNullStr(VerkaufAdoRec!FiveRxXml)
        flxInfo(0).TextMatrix(1, 5) = sAktion
        If (iAktion = 1) Then
            If (XmlAbschnitt(sAktion, "eRezeptStatus") <> "") Then
                sTag = "STATUS"
            Else
                sTag = "VSTATUS"
            End If
        ElseIf (iAktion = 5) Then
            sTag = "STATUS"
        Else
            sTag = "RZLIEFERID"
        End If
'        ind = InStr(UCase(sAktion), "<" + sTag + ">")
'        If (ind > 0) Then
'            sAktion = Mid(sAktion, ind + Len(sTag) + 1 + 1)
'            ind = InStr(UCase(sAktion), "</" + sTag + ">")
'            If (ind > 0) Then
'                sAktion = Left(sAktion, ind - 1)
'                flxInfo(0).TextMatrix(1, 3) = sAktion
'
'                bOk = True
'                If (iAktion = 1) Or (iAktion = 5) Then
'                    bOk = (UCase(sAktion) <> "FEHLER")
'                End If
'                If (bOk) Then
'                    flxInfo(0).TextMatrix(1, 0) = Chr(214)
'                End If
'            End If
'        End If
        'MsgBox (sTag + ". " + sAktion)
        sAktion = XmlAbschnitt(sAktion, sTag)
        If (sAktion <> "") Then
            flxInfo(0).TextMatrix(1, 3) = sAktion
            
            bOk = True
            If (iAktion = 1) Or (iAktion = 5) Then
                bOk = (UCase(sAktion) <> "FEHLER")
            End If
            If (bOk) Then
                flxInfo(0).TextMatrix(1, 0) = Chr(214)
            End If
        End If
    End If
    
'    bOk = True
''    If (UCase(sAktion) = "ABGEBEN") Then
''        bOk = (CheckNullStr(VerkaufAdoRec!DispenseXml) <> "") And (CheckNullStr(VerkaufAdoRec!ResultXml) <> "")
''    ElseIf (UCase(sAktion) = "ABRUF") Then
''        bOk = (CheckNullStr(VerkaufAdoRec!BundleXml) <> "")
''    End If
'
'    h = IIf(bOk, Chr(214), "")
'    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "DD.MM.YY HH:mm")
'    h = h + vbTab + sAktion
'    flxInfo(0).AddItem h, 1
    
    VerkaufAdoRec.MoveNext
Loop
VerkaufAdoRec.Close

With flxInfo(0)
    If (.Rows > .FixedRows) Then
        .FillStyle = flexFillRepeat
        .row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .col
        .CellFontName = "Symbol"
        
        .row = .FixedRows
        .col = 4
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = flexSortStringDescending ' 4   'Zahlen absteigend
        .FillStyle = flexFillSingle
    
        .Enabled = True
        .row = .FixedRows
        .col = 0
        .ColSel = .Cols - 1
        
        .ScrollBars = flexScrollBarVertical
            
        .BackColor = wpara.nlFlexBackColor 'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor  'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
'        If (.SelectionMode = flexSelectionFree) Then
'            .BackColorSel = RGB(135, 61, 52)
'            .ForeColorSel = vbWhite '.ForeColor
'        Else
            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
            .ForeColorSel = .ForeColor
'        End If
        .Appearance = 0
    End If
    
End With

Call DefErrPop
End Sub

Sub ZeigeInfoBereich(pzn$)
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
Dim j%, lief%
Dim h$, h2$, h3$

'With flxInfo(0)
'    .Redraw = False
'    .Rows = .FixedRows
'
'    SQLStr$ = "SELECT * FROM FaktZeilen WHERE BelegId='" + pzn$ + "'"
'    SQLStr$ = SQLStr$ + " ORDER BY Zeile"
'    Set FArtRec = FaktDB.OpenRecordset(SQLStr$)
'    If (FArtRec.RecordCount > 0) Then
'        FArtRec.MoveFirst
'        Do
'            If (FArtRec.EOF) Then
'                Exit Do
'            End If
'
'            .AddItem ""
'            .row = .Rows - 1
'
''            Call ZeigeBelegZeile(flxInfo(0), .row)
'
'            FArtRec.MoveNext
'        Loop
'    End If
'
'    .Redraw = True
'    If (.Rows > .FixedRows) Then
'        .row = .FixedRows
'        .col = 0
'        .ColSel = .Cols - 1
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
Dim i%, gef%, erg%
Dim ch$
        
'txtMatchcode.text = SearchLetter$
'txtMatchcode.SetFocus
'cmdOk(0).Value = True

'erg% = SuchArtikel%(UCase(SearchLetter$), opBereich.ArbeitAnzZeilen)

'gef% = -1
'With flxarbeit(0)
'    For i% = (.row + 1) To (.Rows - 1)
'        ch$ = Left$(.TextMatrix(i%, 4), 1)
'        If (ch$ = SearchLetter$) Then
'            gef% = i%
'            Exit For
'        End If
'    Next i%
'
'    If (gef% < 0) Then
'        For i% = 1 To (.row - 1)
'            ch$ = Left$(.TextMatrix(i%, 4), 1)
'            If (ch$ = SearchLetter$) Then
'                gef% = i%
'                Exit For
'            End If
'        Next i%
'    End If
'
'    If (gef% > 0) Then
''        Call HighlightZeile(True)
'        .row = gef%
'        .col = 0
'        .ColSel = .Cols - 1
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
''        Call HighlightZeile
'        Call FormKurzInfo
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
Dim i%, j%, spBreite%, iColWidth%
Dim sp&
Dim FormStr$
            
With flxarbeit(0)
'    For i% = .Font.Size To 5 Step -1
    For i% = 12 To 5 Step -1
        Font.Size = i%
'        MsgBox (CStr(i) + "  " + CStr(TextWidth(String(150, "X"))) + " " + CStr(.Width))
        If (TextWidth(String(200, "X")) < .Width) Then
            Exit For
        End If
    Next i%
    .Font.Size = Font.Size
    Font.Size = Font.Size + 2
'    Font.Size = 18
    
    For i = 1 To .Cols
        .ColWidth(i - 1) = 0
    Next i
    
'    sFormatStr = "|<AbgabeDat|<TaskId|<RezeptNr|<PrescriptionId|^VO#|^ABG#|<Pzn|<Packung|||<Dar|^N|<ChargenNr|<Dosierung|^AI#|<KK|<KostentraegerName|"
'    'wegen des irrtümlichen Anzeige der unsichtbaren Spalten rechts in PatientName
'    'sFormatStr = sFormatStr + "<ZuzahlungsStatus|^ZuzFr|<Versicherter|<PatientName|<RezeptTyp|<VerordnungsTyp|<Kategorie|<AccessCode|<Secret|"
'    'sFormatStr = sFormatStr + "<Dosieranweisung|^AutIdem|PatientTyp|"
'    sFormatStr = sFormatStr + "<ZuzahlungsStatus|^ZuzFr|<Versicherter|<PatientName||||||"
'    sFormatStr = sFormatStr + "|||"
'    sFormatStr = sFormatStr + "^Impf|^BVG|^Noct|<AbrechBis|<EinreichBis|<BundleKBV|Quittung|<Kz|<VerkaufDat|<DispenseDat|<EinreichDat|>GesBrutto|>Zuzahl|>ALBVVG|<Status||"
     
    Font.Bold = True
    .ColWidth(0) = TextWidth("XX")
    .ColWidth(1) = TextWidth(String(8, "9"))
    .ColWidth(2) = TextWidth(String(18, "9"))
    .ColWidth(3) = TextWidth(String(12, "9"))
    .ColWidth(5) = TextWidth(String(5, "X"))
    .ColWidth(6) = TextWidth(String(5, "X"))
    .ColWidth(7) = TextWidth(String(9, "X"))
    .ColWidth(8) = TextWidth(String(29, "X"))

    .ColWidth(9) = TextWidth(String(3, "X"))
    .ColWidth(10) = TextWidth(String(4, "X"))
    .ColWidth(11) = TextWidth(String(6, "X"))
    .ColWidth(12) = TextWidth(String(3, "X"))
    .ColWidth(13) = TextWidth(String(7, "X"))
    .ColWidth(14) = TextWidth(String(10, "X"))
    
    .ColWidth(15) = TextWidth(String(3, "X"))
    .ColWidth(16) = TextWidth(String(8, "9"))
    .ColWidth(17) = TextWidth(String(15, "X"))
    
    .ColWidth(19) = TextWidth(String(4, "X"))
    .ColWidth(20) = TextWidth(String(9, "9"))
    .ColWidth(21) = TextWidth(String(17, "X"))
    
'    .ColWidth(27) = TextWidth(String(8, "X"))
    .ColWidth(30) = TextWidth(String(3, "X"))
    .ColWidth(31) = TextWidth(String(3, "X"))
    .ColWidth(32) = TextWidth(String(3, "X"))
    
    .ColWidth(33) = TextWidth(String(7, "X"))
    .ColWidth(34) = TextWidth(String(7, "X"))
    
    .ColWidth(37) = TextWidth(String(7, "X"))
    
    .ColWidth(38) = TextWidth(String(12, "X"))
    .ColWidth(39) = TextWidth(String(12, "X"))
    .ColWidth(40) = TextWidth(String(12, "X"))
    .ColWidth(41) = TextWidth(String(12, "X"))
    
    .ColWidth(42) = TextWidth(String(8, "9"))
    .ColWidth(43) = TextWidth(String(8, "9"))
    .ColWidth(44) = TextWidth(String(7, "9"))
    .ColWidth(45) = TextWidth(String(15, "9"))
    
    .ColWidth(47) = wpara.FrmScrollHeight   '+ 2 * wpara.FrmBorderHeight
    Font.Bold = False

'    spBreite% = 0
'    For i% = 1 To .Cols - 1
'        If (.ColWidth(i%) > 0) And (i% < (.Cols - 1)) Then
'            .ColWidth(i%) = .ColWidth(i%) + TextWidth("XX")
'        End If
'        spBreite% = spBreite% + .ColWidth(i%)
'    Next i%
'    If (spBreite% > .Width) Then
'        spBreite% = .Width
'    End If
'    .ColWidth(6) = .Width - spBreite% - 90
End With
    
With flxInfo(0)
    .Font.Size = flxarbeit(0).Font.Size
    Font.Bold = True
    .ColWidth(0) = TextWidth("XX")
    .ColWidth(1) = TextWidth(String(15, "9"))
    .ColWidth(2) = TextWidth(String(20, "X"))
    .ColWidth(3) = 0 ' TextWidth(String(40, "X"))
    .ColWidth(4) = 0    'TextWidth(String(20, "X"))
    
    spBreite% = 0
    For i% = 0 To (.Cols - 1)
        '.ColWidth(i%) = .ColWidth(i%) + TextWidth("XX")
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    iColWidth% = .Width - spBreite% - 90
    If (iColWidth% < 0) Then
        iColWidth% = 0
    End If
    .ColWidth(3) = iColWidth%
    Font.Bold = False
End With

Call DefErrPop: Exit Sub
'With flxarbeit(0)
'    For i% = .Font.Size To 5 Step -1
'        Font.Size = i%
'        If (TextWidth(String(90, "X")) < .Width) Then
'            Exit For
'        End If
'    Next i%
'    .Font.Size = Font.Size
'
'    Font.Bold = True
'
'    .ColWidth(0) = 0
'    .ColWidth(1) = TextWidth(String(2, "X"))
'    .ColWidth(2) = 0    'TextWidth(String(5, "9"))
'    .ColWidth(3) = 0
'    .ColWidth(4) = 15
'    .ColWidth(5) = TextWidth("99.99.9999")
'    .ColWidth(6) = 0    'TextWidth("99:99")
'    .ColWidth(7) = TextWidth(String(5, "9"))
'    .ColWidth(8) = TextWidth(String(8, "X"))
'    .ColWidth(9) = TextWidth(String(5, "9"))
'    .ColWidth(10) = TextWidth(String(7, "9"))
'    .ColWidth(11) = TextWidth("999999.99 ")
'    .ColWidth(12) = TextWidth("999999.99 ")
'    .ColWidth(13) = TextWidth(String(7, "X"))
'    .ColWidth(14) = 0
'    .ColWidth(15) = 0
'    .ColWidth(16) = 0
'    .ColWidth(.Cols - 1) = wpara.FrmScrollHeight
'
'    spBreite% = 0
'    For i% = 0 To (.Cols - 1)
'        If (.ColWidth(i%) = 15) Then
'            .ColWidth(i%) = TextWidth("XX")
'        ElseIf (.ColWidth(i%) > 0) And (i% < (.Cols - 1)) Then
'            .ColWidth(i%) = .ColWidth(i%) + TextWidth("XX")
'        End If
'        spBreite% = spBreite% + .ColWidth(i%)
'    Next i%
'    iColWidth% = .Width - spBreite% - 90
'    If (iColWidth% < 0) Then
'        iColWidth% = 0
'    End If
'    .ColWidth(3) = iColWidth%
'
'    Font.Bold = False
'End With
    
'With flxInfo(0)
'    Font.Bold = True
'    .ColWidth(0) = 0
'    .ColWidth(1) = TextWidth(String(7, "9"))
'    .ColWidth(2) = 0
'    .ColWidth(3) = TextWidth("XXXXXX")
'    .ColWidth(4) = TextWidth("XXX")
'    .ColWidth(5) = TextWidth("99999.99 ")
'    .ColWidth(6) = TextWidth("99999.99 ")
'    .ColWidth(7) = TextWidth("999999.99 ")
'    .ColWidth(8) = TextWidth("99999.99 ")
'    .ColWidth(9) = TextWidth("99999.99 ")
'    .ColWidth(10) = TextWidth("XXXX")
'    .ColWidth(11) = TextWidth("999")
'    .ColWidth(12) = TextWidth("Rezept mit Zuz")
'    .ColWidth(13) = 0
'    .ColWidth(14) = 0
'    .ColWidth(15) = 0
'    .ColWidth(16) = 0
'    .ColWidth(17) = 0
'    .ColWidth(18) = 0
'    .ColWidth(19) = 0
'    .ColWidth(20) = wpara.FrmScrollHeight   '+ 2 * wpara.FrmBorderHeight
'    Font.Bold = False
'
'    spBreite% = 0
'    For i% = 1 To .Cols - 1
'        If (.ColWidth(i%) > 0) Then
'            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
'        End If
'        spBreite% = spBreite% + .ColWidth(i%)
'    Next i%
'    If (spBreite% > .Width) Then
'        spBreite% = .Width
'    End If
'    .ColWidth(2) = .Width - spBreite%
'End With
    
With flxInfo(0)
    sp& = .Width / 6    '8
    .ColWidth(0) = 0    '2 * sp&
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
Dim i%, spBreite%, ydiff%, ydiff2%
Dim wi&

ydiff% = (cboAuswahl(0).Height - lblAuswahl(0).Height) / Screen.TwipsPerPixelY
ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY

With lblAuswahl(0)
    .Left = wpara.LinksX
    .Top = wpara.TitelY
End With
With cboAuswahl(0)
    .Left = lblAuswahl(0).Left + lblAuswahl(0).Width + 150
    .Top = lblAuswahl(0).Top - ydiff%
    .Width = TextWidth(String(15, "X"))
'    .ListIndex = 3
'    .TabIndex = txtAuswahl(1).TabIndex
End With
With cboAuswahlDatum
    .Left = cboAuswahl(0).Left + cboAuswahl(0).Width + 300
    .Top = cboAuswahl(0).Top
    .Width = TextWidth(String(18, "X"))
End With
For i% = 0 To 0
    With cboAuswertung(i%)
        .Left = cboAuswahlDatum.Left + cboAuswahlDatum.Width + 300
        .Top = cboAuswahl(0).Top
        .Width = TextWidth(String(15, "X"))
    End With
    With txtAuswertung(2 * i%)
        .Left = cboAuswertung(i).Left + cboAuswertung(i).Width + 300
        .Top = cboAuswertung(i).Top
    End With
    With txtAuswertung(2 * i% + 1)
        .Left = txtAuswertung(2 * i).Left + txtAuswertung(2 * i).Width + 150
        .Top = cboAuswertung(i).Top
'        If (i = 7) Then
'            iLinksX = .Left + .Width + 900  ' 1500
'        End If
    End With
    With chkAuswertung(i)
        .Left = txtAuswertung(2 * i + 1).Left + txtAuswertung(2 * i + 1).Width + 600
        .Top = cboAuswertung(i).Top
    End With
    
    With txtSuchen
        .Left = chkAuswertung(i).Left + chkAuswertung(i).Width + 1800
        .Top = cboAuswertung(i).Top
        .Width = TextWidth(String(15, "X"))
    End With
    With cmdSuchen
        .Left = txtSuchen.Left + txtSuchen.Width + 90
        
        ydiff2 = (cmdSuchen.Height - lblAuswahl(0).Height) / Screen.TwipsPerPixelY
        ydiff2 = (ydiff2 \ 2) * Screen.TwipsPerPixelY
        .Top = lblAuswahl(0).Top - ydiff2
    End With
Next i%
i = 0
With cboAuswahlRezeptTyp
    .Left = txtAuswertung(2 * i + 1).Left + txtAuswertung(2 * i + 1).Width + 600
    .Top = cboAuswertung(i).Top
    .Width = TextWidth(String(15, "X"))
End With
    

With flxSummen
    .Top = flxarbeit(0).Top + flxarbeit(0).Height
    .Left = flxarbeit(0).Left
    .Height = opBereich.ZeilenHoeheY + 90
    .Width = flxarbeit(0).Width

    .Cols = 10
    .Rows = 1
    .FixedCols = 0
    .FixedRows = 0
    .Rows = 1
    .row = 0
    
'    wi = 0
'    For i% = 0 To 40
'        wi = wi + flxarbeit(0).ColWidth(i%)
'    Next i%
'    .ColWidth(0) = wi

    .Font.Bold = True
    
    wi = 0
    For i = 1 To 8
        .ColWidth(i) = flxarbeit(0).ColWidth(IIf(i Mod 2, 2, 3))
        wi = wi + .ColWidth(i)
    Next
    i = 9
    .ColWidth(i) = 150
    wi = wi + .ColWidth(i)
    
    .ColWidth(0) = .Width - wi
    
'    .ColWidth(2) = flxarbeit(0).ColWidth(42)
'    .ColWidth(3) = flxarbeit(0).ColWidth(flxarbeit(0).Cols - 1)
    
'    wi% = 0
'    For i% = 9 To 13
'        wi% = wi% + flxarbeit(0).ColWidth(i%)
'    Next i%
'    .ColWidth(3) = wi%
    
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColAlignment(8) = flexAlignRightCenter
    
'    .row = 0
'    .col = 1
'    .CellBackColor = vbWhite
'    .CellAlignment = flexAlignRightCenter
'
'    .row = 0
'    .col = 2
'    .CellBackColor = vbWhite
'    .CellAlignment = flexAlignRightCenter
    
    .FillStyle = flexFillRepeat
    .row = 0
    .col = 1
    .ColSel = 2
    .CellBackColor = vbWhite
    .row = 0
    .col = 5
    .ColSel = 6
    .CellBackColor = vbWhite
    .FillStyle = flexFillSingle
    
    .GridLines = flexGridFlat
    
    If (para.Newline) Then
        .BorderStyle = flexBorderNone
        .Width = .Width - 90
        .Height = .Height - 90
        
        .FillStyle = flexFillRepeat
        .row = 0
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellBackColor = RGB(199, 176, 123)
    
        .row = 0
        .col = 1
        .RowSel = .row
        .ColSel = 2
        .CellBackColor = RGB(232, 217, 172)
        
        .row = 0
        .col = 5
        .RowSel = .row
        .ColSel = 6
        .CellBackColor = RGB(232, 217, 172)
        .FillStyle = flexFillSingle
    End If
    
'    .Visible = False
End With

'On Error Resume Next
'
'With lblAnzahlWert(0)
'    .Left = Me.ScaleWidth - wpara.LinksX - .Width
'    .Top = lblMatchcode.Top - 45
'End With
'For i% = 1 To (AnzLblAnzahl% - 1)
'    With lblAnzahlWert(i%)
'        .Left = lblAnzahlWert(i% - 1).Left
'        .Top = lblAnzahlWert(i% - 1).Top + lblAnzahlWert(i% - 1).Height
'    End With
'Next i%
'
'With lblAnzahl(0)
'    .Left = lblAnzahlWert(0).Left - .Width - 150
'    .Top = lblMatchcode.Top - 45
'End With
'For i% = 1 To (AnzLblAnzahl% - 1)
'    With lblAnzahl(i%)
'        .Left = lblAnzahl(i% - 1).Left
'        .Top = lblAnzahl(i% - 1).Top + lblAnzahl(i% - 1).Height
'    End With
'Next i%

If (para.Newline) Then
    With cboAuswahl(0)
'        .Appearance = 0
        .BackColor = vbWhite
        Call wpara.ControlBorderless(cboAuswahl(0))
    End With
    With picBack(0)
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
    End With
    With cboAuswahl(0)
        RoundRect picBack(0).hdc, (.Left - 60) / Screen.TwipsPerPixelX, (.Top - 30) / Screen.TwipsPerPixelY, (.Left + .Width + 60) / Screen.TwipsPerPixelX, (.Top + .Height + 15) / Screen.TwipsPerPixelY, 10, 10
    End With
    
    With cboAuswahlDatum
'        .Appearance = 0
        .BackColor = vbWhite
        Call wpara.ControlBorderless(cboAuswahlDatum)
    End With
    With cboAuswahlDatum
        RoundRect picBack(0).hdc, (.Left - 60) / Screen.TwipsPerPixelX, (.Top - 30) / Screen.TwipsPerPixelY, (.Left + .Width + 60) / Screen.TwipsPerPixelX, (.Top + .Height + 15) / Screen.TwipsPerPixelY, 10, 10
    End With
    
    With cboAuswertung(0)
'        .Appearance = 0
        .BackColor = vbWhite
        Call wpara.ControlBorderless(cboAuswertung(0))
    End With
    With cboAuswertung(0)
        RoundRect picBack(0).hdc, (.Left - 60) / Screen.TwipsPerPixelX, (.Top - 30) / Screen.TwipsPerPixelY, (.Left + .Width + 60) / Screen.TwipsPerPixelX, (.Top + .Height + 15) / Screen.TwipsPerPixelY, 10, 10
    End With
    With cboAuswahlRezeptTyp
'        .Appearance = 0
        .BackColor = vbWhite
        Call wpara.ControlBorderless(cboAuswahlRezeptTyp)
    End With
    With cboAuswahlRezeptTyp
        RoundRect picBack(0).hdc, (.Left - 60) / Screen.TwipsPerPixelX, (.Top - 30) / Screen.TwipsPerPixelY, (.Left + .Width + 60) / Screen.TwipsPerPixelX, (.Top + .Height + 15) / Screen.TwipsPerPixelY, 10, 10
    End With
    
    
    For i = 0 To 1
        With txtAuswertung(i)
            If (.Appearance = 1) Then
                Call wpara.ControlBorderless(txtAuswertung(i), 2, 2)
            Else
                Call wpara.ControlBorderless(txtAuswertung(i), 1, 1)
            End If
        End With
    Next
        
    
    With txtAuswertung(0)
        .BackColor = vbWhite
        With .Container
            .ForeColor = RGB(180, 180, 180) ' vbWhite
            .FillStyle = vbSolid
            .FillColor = vbWhite

            RoundRect .hdc, (txtAuswertung(0).Left - 60) / Screen.TwipsPerPixelX, (txtAuswertung(0).Top - 30) / Screen.TwipsPerPixelY, (txtAuswertung(0).Left + txtAuswertung(0).Width + 60) / Screen.TwipsPerPixelX, (txtAuswertung(0).Top + txtAuswertung(0).Height + 15) / Screen.TwipsPerPixelY, 10, 10
        End With
    End With
    
    With lblchkAuswertung(chkAuswertung(0).index)
        .BackStyle = 0 'duchsichtig
        .Caption = chkAuswertung(0).Caption
        .Left = chkAuswertung(0).Left + chkAuswertung(0).Width + 60
        .Top = chkAuswertung(0).Top
        .Width = TextWidth(.Caption) + 90
        .TabIndex = chkAuswertung(0).TabIndex
        .Visible = True
    End With
    
    With txtSuchen
        If (.Appearance = 1) Then
            Call wpara.ControlBorderless(txtSuchen, 2, 2)
        Else
            Call wpara.ControlBorderless(txtSuchen, 1, 1)
        End If
    End With
    With txtSuchen
        .BackColor = vbWhite
        With .Container
            .ForeColor = RGB(180, 180, 180) ' vbWhite
            .FillStyle = vbSolid
            .FillColor = vbWhite

            RoundRect .hdc, (txtSuchen.Left - 60) / Screen.TwipsPerPixelX, (txtSuchen.Top - 30) / Screen.TwipsPerPixelY, (txtSuchen.Left + txtSuchen.Width + 60) / Screen.TwipsPerPixelX, (txtSuchen.Top + txtSuchen.Height + 15) / Screen.TwipsPerPixelY, 10, 10
        End With
    End With

End If

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

lblAuswahl(0).BackColor = wpara.FarbeArbeit

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
Dim i%, j%, ind%
Dim l&
Dim h$, h2$, Key$

'h$ = "N"
'l& = GetPrivateProfileString(INI_SECTION, "ArtikelStatistik", "N", h$, 2, INI_DATEI)
'h$ = Left$(h$, l&)
'If (h$ = "J") Then
'    ArtikelStatistik% = True
'Else
'    ArtikelStatistik% = False
'End If
'
'h$ = "080"
'l& = GetPrivateProfileString(INI_SECTION, "ProzentA", h$, h$, 4, INI_DATEI)
'h$ = Left$(h$, l&)
'ABCprozentA% = xVal(h$)
'
'h$ = "095"
'l& = GetPrivateProfileString(INI_SECTION, "ProzentB", h$, h$, 4, INI_DATEI)
'h$ = Left$(h$, l&)
'ABCprozentB% = xVal(h$)
'
'h$ = Space$(4)
'l& = GetPrivateProfileString(INI_SECTION, "DruckBisProzent", h$, h$, 5, INI_DATEI)
'h$ = Trim$(Left$(h$, l&))
'ABCdruckProzent% = xVal(h$)
'
'h$ = Space$(4)
'l& = GetPrivateProfileString(INI_SECTION, "DruckBisPositionen", h$, h$, 5, INI_DATEI)
'h$ = Trim$(Left$(h$, l&))
'ABCdruckPositionen% = xVal(h$)
'
'For i% = 0 To 2
'    h$ = Space$(20)
'    key$ = "Darstellung" + Format(i%, "0")
'    l& = GetPrivateProfileString(INI_SECTION, key$, h$, h$, 21, INI_DATEI)
'    h$ = Left$(h$, l&)
'
'    If (para.Newline) Then
'        ABCdarstellung&(i%, 0) = vbBlack
'        ABCdarstellung&(i%, 1) = vbWhite
'    Else
'        ABCdarstellung&(i%, 0) = flxarbeit(0).ForeColor
'        ABCdarstellung&(i%, 1) = flxarbeit(0).BackColor
'    End If
''    If (i% = 0) Then
''        ABCdarstellung&(i%, 1) = flxarbeit(0).BackColor
''    ElseIf (i% = 1) Then
''        ABCdarstellung&(i%, 1) = RGB(150, 150, 150)
''    Else
''        ABCdarstellung&(i%, 1) = RGB(100, 100, 100)
''    End If
'
'    If (Trim(h$) <> "") Then
'        ind% = InStr(h$, ",")
'        If (ind% > 0) Then
'            ABCdarstellung&(i%, 1) = wpara.BerechneFarbWert(Mid$(h$, ind% + 1))
'            h$ = Left$(h$, ind% - 1)
'        End If
'        If (Trim(h$) <> "") Then
'            ABCdarstellung&(i%, 0) = wpara.BerechneFarbWert(h$)
'        End If
'    End If
'Next i%

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
Dim i%
Dim l&
Dim h$, Key$

'l& = WritePrivateProfileString(INI_SECTION, "ProzentA", Str$(ABCprozentA%), INI_DATEI)
'l& = WritePrivateProfileString(INI_SECTION, "ProzentB", Str$(ABCprozentB%), INI_DATEI)
'l& = WritePrivateProfileString(INI_SECTION, "DruckBisProzent", Str$(ABCdruckProzent%), INI_DATEI)
'l& = WritePrivateProfileString(INI_SECTION, "DruckBisPositionen", Str$(ABCdruckPositionen%), INI_DATEI)
'
'For i% = 0 To 2
'    h$ = ""
'    If (ABCdarstellung&(i%, 0) <> flxarbeit(0).ForeColor) Then
'        h$ = Hex$(ABCdarstellung&(i%, 0))
'    End If
'    If (ABCdarstellung&(i%, 1) <> flxarbeit(0).BackColor) Then
'        h$ = h$ + "," + Hex$(ABCdarstellung&(i%, 1))
'    End If
'
'    Key$ = "Darstellung" + Format(i%, "0")
'    l& = WritePrivateProfileString(INI_SECTION, Key$, h$, INI_DATEI)
'Next i%
       
Call DefErrPop
End Sub

Private Sub flxarbeit_KeyPress(index As Integer, KeyAscii As Integer)
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
        Call ToggleAuswahlZeile
        Call NaechsteAuswahlZeile
    ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
        Call SelectZeile(UCase(Chr$(KeyAscii)))
    End If
End If

Call DefErrPop
End Sub

Sub NaechsteAuswahlZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NaechsteAuswahlZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    .col = 0
    .ColSel = .Cols - 1
End With

'Call AuswahlRowChange(vbKeyDown)

Call DefErrPop
End Sub

'Sub HighlightZeile(Optional NurNormalMachen% = False)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("HighlightZeile")
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
'Dim aRow%, aCol%, rInd%, ZeilenWechsel%
'Dim RkLaufNr&
'Dim h$, KalkText$, DirektWerte$
'Static aRkLaufNr&, aFlexRow%
'
'ZeilenWechsel% = False
'
'With flxarbeit(0)
'    If (NurNormalMachen%) Then
''        .HighLight = flexHighlightNever
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
'        aRkLaufNr& = -1&
'        aFlexRow% = .row
''        .HighLight = flexHighlightWithFocus
'        KeinRowColChange% = False
'    Else
'        RkLaufNr& = val(.TextMatrix(.row, 3))
'        If (RkLaufNr& <> aRkLaufNr&) Or (aFlexRow% <> .row) Then
'
''            .HighLight = flexHighlightNever
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
'            Call FormKurzInfo
'
'            aRkLaufNr& = RkLaufNr&
'            aFlexRow% = .row
''            .HighLight = flexHighlightWithFocus
'
'            KeinRowColChange% = False
'
'            ZeilenWechsel% = True
'        End If
'
'    End If
'
''    .col = 0
''    .ColSel = .Cols - 1
'
'End With
'
''Call AnzeigeKommentar
'
'Call DefErrPop
'End Sub

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

Sub AuswahlBefüllen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlBefüllen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ind%, ind2%, row%, zOk%, WhereDa%, ATCCODES_CSV%, iOk%, iAnzRezepte%
Dim iVerordnungsTyp%
Dim SollKuNr&, iKuNr&, AltKuNr&
Dim dGesBrutto#, dGesZuzahlungen#, dGesALBVVG#
Dim h$, h2$, sKurz$, sInfo$, SQLStr2$, sOp$, sTxt$(1), sKz$
Dim iDat As Date
Dim iRec As Recordset
Dim eRezept As New TI_Back.eRezept_OP

'Call ZerlegeKundenNummern(ParamKunden$)

MousePointer = vbHourglass

h$ = cboAuswahl(0).text
SollStatus = Val(Trim(Right(h, 3)))

mnuKomfortSignatur.Enabled = (InStr(h, "(HBA") > 0)

With flxarbeit(0)
    .Redraw = False
    .Rows = .FixedRows
End With


SQLStr = "SELECT DISTINCT TI.*, VK.RezeptNr, VK.eRezTaskID, VK.Datum AS VkDatum FROM TI_eRezepte AS TI"
SQLStr = SQLStr + " LEFT JOIN Verkauf AS VK ON TI.TaskId=VK.eRezTaskID"
SQLStr = SQLStr + " AND VK.Datum = ( SELECT MAX(Datum) FROM Verkauf AS VK2 WHERE VK2.eRezTaskId=VK.eRezTaskId AND LEN(RezeptNr)>0 )"
If (SollStatus < 0) Then
    SQLStr = SQLStr + " WHERE ((za_1_schluessel=-1) or (AbgabeDatum is null) or (RezeptNr is null))"
Else
    SQLStr = SQLStr + " WHERE (OpStatus" + IIf(SollStatus = 0, ">=0", "=" + CStr(SollStatus)) + ")"
    If (SollStatus > 0) Then
        SQLStr = SQLStr + " AND (ZA_1_Schluessel>=0)"
    End If
    If (SollStatus = 3) Then
        SQLStr = SQLStr + " AND (QES" + IIf(InStr(h, "(HBA") > 0, ">", "=") + "0)"
    End If
End If

''''''''''
Dim sAuwahlDatum$
With cboAuswahlDatum
    ind = .ListIndex
    sAuwahlDatum = "AbgabeDatum"
    If (.ListIndex = 1) Then
        sAuwahlDatum = "DispensierDatum"
    ElseIf (.ListIndex = 2) Then
        sAuwahlDatum = "EinreichDatum"
    ElseIf (.ListIndex = 3) Then
        sAuwahlDatum = "AnlageDatum"
    End If
End With

If (ind <= 3) Then
    For i = 0 To 0
        SQLStr2 = ""
        With cboAuswertung(i)
            iOk = 0
            If (i = 5) Then
            ElseIf (.ListIndex > 0) Then
                iOk = True
                
                sOp = Trim(UCase(.text))
                sTxt(0) = ""
                sTxt(1) = ""
                If (txtAuswertung(2 * i).Visible) Then
                    sTxt(0) = Trim(UCase(txtAuswertung(2 * i)))
                End If
                If (txtAuswertung(2 * i + 1).Visible) Then
                    sTxt(1) = Trim(UCase(txtAuswertung(2 * i + 1)))
                End If
                
                iOk = (Left(sOp, 5) <> String(5, "-"))
                If (iOk) Then
    '                If (i = 0) Or (i = 1) Or (i = 3) Or (i = 8) Or (i = 9) Then
    '                Else
    '                    iOk = (sTxt(0) <> "")
    '                    If (iOk) And (sOp = "ZWISCHEN") Then
    '                        iOk = (sTxt(1) <> "")
    '                    End If
    '                End If
                    If (txtAuswertung(2 * i).Visible) Then
                        iOk = (sTxt(0) <> "")
                    End If
                    If (iOk) And (txtAuswertung(2 * i + 1).Visible) Then
                        iOk = (sTxt(0) <> "")
                    End If
                End If
                
                If (iOk) Then
                    iDat = Left(sTxt(0), 2) + "." + Mid(sTxt(0), 3, 2) + ".20" + Mid(sTxt(0), 5, 2)
        '                    If (sOp = "ZWISCHEN") Then
                    If (.text = "=") Then
                        SQLStr2 = sAuwahlDatum + ">='" + Format(iDat, "DD.MM.YYYY") + "'"
                        SQLStr2 = SQLStr2 + " AND " + sAuwahlDatum + "<'" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                    ElseIf (txtAuswertung(2 * i + 1).Visible) Then
                        SQLStr2 = sAuwahlDatum + ">='" + Format(iDat, "DD.MM.YYYY") + "'"
                        iDat = Left(sTxt(1), 2) + "." + Mid(sTxt(1), 3, 2) + ".20" + Mid(sTxt(1), 5, 2)
                        SQLStr2 = SQLStr2 + " AND " + sAuwahlDatum + "<'" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                    Else
                        If (.text = "<") Then
                            SQLStr2 = sAuwahlDatum + "<='" + Format(iDat, "DD.MM.YYYY") + "'"
                        ElseIf (.text = "<=") Then
                            SQLStr2 = sAuwahlDatum + "<'" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                        ElseIf (.text = "=") Then
                            SQLStr2 = sAuwahlDatum + ">'" + Format(iDat - 1, "DD.MM.YYYY") + "'"
                            SQLStr2 = SQLStr2 + " AND " + sAuwahlDatum + "<'" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                        ElseIf (.text = "<>") Then
                            SQLStr2 = "(" + sAuwahlDatum + ">='" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                            SQLStr2 = SQLStr2 + " OR " + sAuwahlDatum + "<'" + Format(iDat, "DD.MM.YYYY") + "')"
                        ElseIf (.text = ">=") Then
                            SQLStr2 = sAuwahlDatum + ">'" + Format(iDat - 1, "DD.MM.YYYY") + "'"
                        ElseIf (.text = ">") Then
                            SQLStr2 = sAuwahlDatum + ">='" + Format(iDat + 1, "DD.MM.YYYY") + "'"
                        End If
                    End If
                End If
            End If
        End With
'        If (chkAuswertung(0).Value) Then
'    '        SQLStr2 = SQLStr2 + " AND VerordnungsTyp=" + CStr(ti_back.VerordnungsTyp_OP_Rezepturverordnung)
'            SQLStr2 = SQLStr2 + " AND ((VerordnungsTyp=" + CStr(TI_Back.VerordnungsTyp_OP_Rezepturverordnung) + ") OR (RezepturHerstellungen<>'') OR (TI.Pzn='09999011'))"
'        End If
        ind2 = cboAuswahlRezeptTyp.ListIndex
        If (ind2 = 1) Then
            SQLStr2 = SQLStr2 + " AND ((VerordnungsTyp=" + CStr(TI_Back.VerordnungsTyp_OP_Rezepturverordnung) + ") OR (RezepturHerstellungen<>'') OR (TI.Pzn='09999011'))"
        ElseIf (ind2 = 2) Then
            SQLStr2 = SQLStr2 + " AND (RezeptTypId=910)"
        End If
        If (SQLStr2 <> "") Then
            SQLStr = SQLStr + " AND "
            SQLStr = SQLStr + SQLStr2
        End If
    Next i
Else
    With cboAuswahlDatum
        h = .text
    End With
    
    Select Case h
'        Case "ChargenNr"
        Case "KK"
            h = "KostentraegerIK"
'        Case "KostentraegerName"
        Case "Packung"
            h = "Verordnungstext"
'        Case "PatientName"
'        Case "Pzn"
'        Case "RezeptNr"
'        Case "TaskId"
        Case "Versicherter"
            h = "PatientWert"
    End Select
    
    SQLStr2 = h + " LIKE '%" + txtAuswertung(0).text + "%'"
    If (SQLStr2 <> "") Then
        SQLStr = SQLStr + " AND "
        SQLStr = SQLStr + SQLStr2
    End If
End If

''''''''''
iAnzRezepte = 0
dGesBrutto = 0
dGesZuzahlungen = 0
dGesALBVVG = 0

'SQLStr = SQLStr + " ORDER BY AnlageDatum DESC"
SQLStr = SQLStr + " ORDER BY AbgabeDatum DESC"
'MsgBox (SQLStr)
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'If (FabsErrf <> 0) Then
'    Call iMsgBox("keine passenden Rezepte gespeichert !")
'    Call DefErrPop: Exit Sub
'End If
Do
    If (VerkaufAdoRec.EOF) Then
        Exit Do
    End If
    
    h = ""
'    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "DD.MM.YY HH:mm")
1    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AbgabeDatum), "DD.MM.YY")
'    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AbgabeDatum), "DD.MM.YY HH:mm")
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!TaskId)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!RezeptNr)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!PrescriptionID)
    h = h + vbTab + CStr(CheckNullLong(VerkaufAdoRec!AnzahlVerordnetePackungen))
    h = h + vbTab + CStr(CheckNullLong(VerkaufAdoRec!AnzahlAbgegebenePackungen))
    
    iVerordnungsTyp = CheckNullLong(VerkaufAdoRec!VerordnungsTyp)
    If (iVerordnungsTyp = VerordnungsTyp_OP_Freitextverordnung) Then
        h = h + vbTab + "FT-VO"
        h = h + vbTab + CheckNullStr(VerkaufAdoRec!Verordnungstext)
    ElseIf (iVerordnungsTyp = VerordnungsTyp_OP_Rezepturverordnung) Then
        h = h + vbTab + "Rezeptur-VO"
        h = h + vbTab + CheckNullStr(VerkaufAdoRec!Verordnungstext)
    ElseIf (iVerordnungsTyp = VerordnungsTyp_OP_Wirkstoffverordnung) Then
        h = h + vbTab + "W-VO"
        h = h + vbTab + CheckNullStr(VerkaufAdoRec!Wirkstoffname) + " " + CStr(CheckNullDouble(VerkaufAdoRec!Wirkstaerke)) + " " + CheckNullStr(VerkaufAdoRec!WirkstaerkenEinheit)
    Else
        h = h + vbTab + CheckNullStr(VerkaufAdoRec!pzn)
        h = h + vbTab + CheckNullStr(VerkaufAdoRec!Verordnungstext)
    End If
    
    h = h + vbTab + CStr(CheckNullDouble(VerkaufAdoRec!Packungsgroesse))
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Einheit)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Darreichungsform)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Normgroesse)
    
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!ChargenNr)
    
    h2 = CheckNullStr(VerkaufAdoRec!Dosieranweisung)
    If (Len(h2) > 10) Then
        h2 = Left(h2, 10) + " ..."
    End If
    h = h + vbTab + h2
                        
    h = h + vbTab + IIf(CheckNullByte(VerkaufAdoRec!AustauschErlaubt) = 1, "", Chr(214))
                        
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!KostentraegerIK)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!KostentraegerName)
                       
    h = h + vbTab + CStr(CheckNullLong(VerkaufAdoRec!ZuzahlungsStatus))
'    h = h + vbTab + IIf(CheckNullLong(VerkaufAdoRec!ZuzahlungsStatus) = Zuzahlungsstatus_gebuehrenfrei, Chr(214), "")
    h = h + vbTab + IIf(CheckNullLong(VerkaufAdoRec!ZuzahlungsStatus) = 1, Chr(214), "")
                       
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!PatientWert)

'    h = h + vbTab + CheckNullStr(VerkaufAdoRec!PatientName)
    h2 = ""
    Call eRezept.FhirPatient(CheckNullStr(VerkaufAdoRec!Bundle))
    With eRezept.Patient.Name
        h2 = .Nachname_ohne_Vor_und_Zusatz + ", " + .Vorname
    End With
    h = h + vbTab + h2
    
    
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!RezeptTyp)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!VerordnungsTyp_Text)
    
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Kategorie_Text)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!AccessCode)
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Secret)
    
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!Dosieranweisung)
    h = h + vbTab + IIf(CheckNullByte(VerkaufAdoRec!AustauschErlaubt) = 1, Chr(214), "")
    h = h + vbTab + CheckNullStr(VerkaufAdoRec!PatientTyp)
    
    h = h + vbTab + IIf(CheckNullByte(VerkaufAdoRec!Impfstoff) = 1, Chr(214), "")
    h = h + vbTab + IIf(CheckNullByte(VerkaufAdoRec!BVG) = 1, Chr(214), "")
    h = h + vbTab + IIf(CheckNullByte(VerkaufAdoRec!Noctu) = 1, Chr(214), "")
    
    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!FristAbrechnung), "DD.MM.YY")
    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!FristEinreichung), "DD.MM.YY")

    h2 = CheckNullStr(VerkaufAdoRec!Bundle)
    h2 = Replace(h2, Format(vbTab, "0"), " ")
    h = h + vbTab + h2  ' CheckNullStr(VerkaufAdoRec!Bundle)
    h = h + vbTab
    
    sKz = IIf(CheckNullByte(VerkaufAdoRec!Mehrfach) = 1, "MVO", "")
    If (CheckNullByte(VerkaufAdoRec!Unfall) = 1) Then
        sKz = sKz & IIf(sKz <> "", ", ", "") + "UI"
    End If
    If (CheckNullStr(VerkaufAdoRec!Kostentraeger_Typ) <> "GKV") Then
        sKz = sKz & IIf(sKz <> "", ", ", "") + CheckNullStr(VerkaufAdoRec!Kostentraeger_Typ)
    End If
    h = h + vbTab + sKz
    
    iDat = CheckNullDate(VerkaufAdoRec!AnlageDatum)
    h = h + vbTab + IIf(Year(iDat) > 2000, Format(iDat, "DD.MM.YY HH:mm"), "")
    iDat = CheckNullDate(VerkaufAdoRec!VkDatum)
    h = h + vbTab + IIf(Year(iDat) > 2000, Format(iDat, "DD.MM.YY HH:mm"), "")
    iDat = CheckNullDate(VerkaufAdoRec!DispensierDatum)
    h = h + vbTab + IIf(Year(iDat) > 2000, Format(iDat, "DD.MM.YY HH:mm"), "")
    iDat = CheckNullDate(VerkaufAdoRec!EinreichDatum)
    h = h + vbTab + IIf(Year(iDat) > 2000, Format(iDat, "DD.MM.YY HH:mm"), "")
    
    h = h + vbTab + Format(CheckNullDouble(VerkaufAdoRec!GesamtBrutto), "# ##0.00")
    h = h + vbTab + Format(CheckNullDouble(VerkaufAdoRec!GesamtZuzahlung), "# ##0.00")
            
    h = h + vbTab
    If (CheckNullInt(VerkaufAdoRec!ZA_2_Schluessel) = 2) Or (CheckNullInt(VerkaufAdoRec!ZA_2_Schluessel) = 3) Then
        h = h + Format(0.6, "# ##0.00")
        dGesALBVVG = dGesALBVVG + 0.6
    End If
                
    h = h + vbTab + OpStatusStr(CheckNullInt(VerkaufAdoRec!OpStatus))
    
    iAnzRezepte = iAnzRezepte + 1
    dGesBrutto = dGesBrutto + CheckNullDouble(VerkaufAdoRec!GesamtBrutto)
    dGesZuzahlungen = dGesZuzahlungen + CheckNullDouble(VerkaufAdoRec!GesamtZuzahlung)
    
    With flxarbeit(0)
        .AddItem h
        If (CheckNullByte(VerkaufAdoRec!QES) > 0) Then
            .FillStyle = flexFillRepeat
            .row = .Rows - 1
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellBackColor = vbRed
            .FillStyle = flexFillSingle
        End If
        If (InStr(sKz, "SEL") > 0) Then
            .FillStyle = flexFillRepeat
            .row = .Rows - 1
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellBackColor = vbGreen
            .FillStyle = flexFillSingle
        End If
    End With
    
    VerkaufAdoRec.MoveNext
Loop
VerkaufAdoRec.Close


With flxarbeit(0)
    If (.Rows = .FixedRows) Then
        .AddItem vbTab + vbTab + "(leere Liste)"
    End If
End With

'Call AuswahlSortieren

With flxarbeit(0)
    .FillStyle = flexFillRepeat
    
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 15
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 19
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 28
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 30
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 31
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .row = 1
    .col = 32
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    
    .row = 1
    .col = 3
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellBackColor = RGB(200, 200, 200)
    
    .row = 1
    .col = 13
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellBackColor = RGB(200, 200, 200)
    
    .FillStyle = flexFillSingle
    
    .Redraw = True

    If (.Visible) Then
        .row = .FixedRows
        .col = 0
        .ColSel = .Cols - 1
        .SetFocus
    End If
    
    .ScrollBars = flexScrollBarBoth
End With

With flxSummen
    h$ = ""
'    If (GesAnzPackungen! <> 0) Then
'        If (GesAnzPackungen! = CLng(GesAnzPackungen!)) Then
'            h$ = h$ + Format(GesAnzPackungen!, "0")
'        ElseIf (FNX(GesAnzPackungen! * 10) = CLng(GesAnzPackungen! * 10)) Then
'            h$ = h$ + Format(GesAnzPackungen!, "0.0")
'        Else
'            h$ = h$ + Format(GesAnzPackungen!, "0.00")
'        End If
'        h$ = h$ + " Einheit(en)  /  "
'    End If
'    .TextMatrix(0, 0) = h$ + Format(GesAnzPositionen%, "0") + " ArtikelPosition(en)"
'
'    .TextMatrix(0, 1) = Format(GesWertExkl#, "0.00")
'    .TextMatrix(0, 2) = Format(GesWertInkl#, "0.00")
    .TextMatrix(0, 1) = "Anzahl eRezepte"
    .TextMatrix(0, 2) = Format(iAnzRezepte, "0")
    .TextMatrix(0, 3) = "Gesamt-Brutto"
    .TextMatrix(0, 4) = Format(dGesBrutto, "# ### ##0.00")
    .TextMatrix(0, 5) = "Zuzahlungen"
    .TextMatrix(0, 6) = Format(dGesZuzahlungen, "# ### ##0.00")
    .TextMatrix(0, 7) = "ALBVVG"
    .TextMatrix(0, 8) = Format(dGesALBVVG, "# ### ##0.00")
    
    .Visible = (ShowSummen) And (cboAuswahl(0).ListIndex >= 4)
End With

mnuBearbeitenInd(MENU_F5).Enabled = (SollStatus = 2) Or (SollStatus = 4) Or (SollStatus = 5)
mnuBearbeitenInd(MENU_SF5).Enabled = (SollStatus = 2)
mnuBearbeitenInd(MENU_SF6).Enabled = (SollStatus < 6)

For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%

Call opToolbar.ShowToolbar

MousePointer = vbNormal

Call FormKurzInfo

Call DefErrPop
End Sub

Sub ZeigeBezügeBelegZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBezügeBelegZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, row%
Dim fAnzPackungen!
Dim dWert#, dWertInkl#, dEK#
Dim h$, h2$, pzn$
Dim sRec As New ADODB.Recordset

'With flxarbeit(0)
'    If (KunBezugSQL) Then
'        pzn$ = PznString(FArtAdoRec!pzn)
'        fAnzPackungen! = FArtAdoRec!Anzahl
'        dEK# = FArtAdoRec!Preis
'        dWert# = fAnzPackungen! * FArtAdoRec!Preis
'        dWertInkl# = dWert '* 1.19
'
'        row% = 0
'        For i% = .FixedRows To (.Rows - 1)
'            If (pzn$ = .TextMatrix(i%, 0)) Then
'                row% = i%
'                If (Trim(pzn$) = "") Or (Val(pzn$) = 9999999) Then
'                    If (FArtAdoRec!text <> .TextMatrix(i%, 2)) Then
'                        row% = 0
'                    ElseIf (FArtAdoRec!menge <> .TextMatrix(i%, 3)) Then
'                        row% = 0
'                    ElseIf (FArtAdoRec!einheit <> .TextMatrix(i%, 4)) Then
'                        row% = 0
'                    End If
'                End If
'                If (row%) Then
'                    Exit For
'                End If
'            End If
'        Next i%
'
'
'        If (row% = 0) Then
'            .AddItem ""
'            row% = .Rows - 1
'
'            .TextMatrix(row%, 0) = pzn$
'
'            .TextMatrix(row%, 2) = FArtAdoRec!text
'            .TextMatrix(row%, 3) = FArtAdoRec!menge
'            .TextMatrix(row%, 4) = FArtAdoRec!einheit
'
'            h = ""
'            SQLStr = "SELECT * FROM Artikel WHERE Pzn=" + pzn
'            sRec.open SQLStr, Artikel.ActiveConn
'            If Not (sRec.EOF) Then
'                h = Format(CheckNullDouble(sRec!EK), "0.00")
'            Else
'                sRec.Close
'                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + SqlPzn(pzn$)
'                sRec.open SQLStr, taxeAdoDB.ActiveConn
'                If Not (sRec.EOF) Then
'                    h = Format(CheckNullDouble(sRec!EK) / 100#, "0.00")
'                End If
'            End If
'            sRec.Close
'            .TextMatrix(row%, 5) = h$
'
'            h = CheckNullStr(FArtAdoRec!ATCCode)
'            .TextMatrix(row%, 6) = h
'            If (h = "") Then
'                SQLStr = "SELECT * FROM Selbstangelegte WHERE Pzn=" + pzn
'                sRec.open SQLStr, Artikel.ActiveConn
'                If Not (sRec.EOF) Then
'                    h = CheckNullStr(sRec!ATCCode)
'                    .TextMatrix(row%, 6) = h
'                End If
'                sRec.Close
'            End If
'            If (h <> "") Then
'                If (Dir(CurDir + "\" + "atccodes.csv") <> "") Then
'                    For i = 1 To UBound(AtcCodes)
'                        If (AtcCodes(i) = h) Then
'                            .TextMatrix(row, 6) = AtcNamen(i)
'                            Exit For
'                        End If
'                    Next i
'                End If
'            End If
'
'            h$ = .TextMatrix(row%, 0)
'            For i% = 2 To 4
'                h$ = h$ + .TextMatrix(row%, i%)
'            Next i%
'            .TextMatrix(row%, 12) = h$
'        Else
'            fAnzPackungen! = xVal(.TextMatrix(row%, 1)) + fAnzPackungen!
'            dWert# = xVal(.TextMatrix(row%, 7)) + dWert
'            dWertInkl# = xVal(.TextMatrix(row%, 8)) + dWertInkl
'        End If
'    Else
'        pzn$ = PznString(FArtRec!pzn)
'        fAnzPackungen! = FArtRec!Anzahl
'        dEK# = CheckNullDouble(FArtRec!Preis)
'        dWert# = fAnzPackungen! * CheckNullDouble(FArtRec!Preis)
'        dWertInkl# = dWert '* 1.19
'
'        row% = 0
'        For i% = .FixedRows To (.Rows - 1)
'            If (pzn$ = .TextMatrix(i%, 0)) Then
'                row% = i%
'                If (Trim(pzn$) = "") Or (Val(pzn$) = 9999999) Then
'                    If (FArtRec!text <> .TextMatrix(i%, 2)) Then
'                        row% = 0
'                    ElseIf (FArtRec!menge <> .TextMatrix(i%, 3)) Then
'                        row% = 0
'                    ElseIf (FArtRec!einheit <> .TextMatrix(i%, 4)) Then
'                        row% = 0
'                    End If
'                End If
'                If (row%) Then
'                    Exit For
'                End If
'            End If
'        Next i%
'
'
'        If (row% = 0) Then
'            .AddItem ""
'            row% = .Rows - 1
'
'            .TextMatrix(row%, 0) = pzn$
'
'            .TextMatrix(row%, 2) = FArtRec!text
'            .TextMatrix(row%, 3) = FArtRec!menge
'            .TextMatrix(row%, 4) = FArtRec!einheit
'
'            h = CheckNullStr(FArtRec!ATCCode)
'            .TextMatrix(row%, 6) = h
'            If (h = "") Then
'                SQLStr = "SELECT * FROM Selbstangelegte WHERE Pzn=" + pzn
'                sRec.open SQLStr, Artikel.ActiveConn
'                If Not (sRec.EOF) Then
'                    h = CheckNullStr(sRec!ATCCode)
'                    .TextMatrix(row%, 6) = h
'                End If
'                sRec.Close
'            End If
'            If (h <> "") Then
'                If (Dir(CurDir + "\" + "atccodes.csv") <> "") Then
'                    For i = 1 To UBound(AtcCodes)
'                        If (AtcCodes(i) = h) Then
'                            .TextMatrix(row, 6) = AtcNamen(i)
'                            Exit For
'                        End If
'                    Next i
'                End If
'            End If
'
'
'            h$ = .TextMatrix(row%, 0)
'            For i% = 2 To 4
'                h$ = h$ + .TextMatrix(row%, i%)
'            Next i%
'            .TextMatrix(row%, 12) = h$
'        Else
'            fAnzPackungen! = xVal(.TextMatrix(row%, 1)) + fAnzPackungen!
'            dWert# = xVal(.TextMatrix(row%, 7)) + dWert
'            dWertInkl# = xVal(.TextMatrix(row%, 8)) + dWertInkl
'        End If
'    End If
'
'    h$ = ""
'    If (fAnzPackungen! <> 0) Then
'        If (fAnzPackungen! = CLng(fAnzPackungen!)) Then
'            h$ = Format(fAnzPackungen!, "0")
'        ElseIf (FNX(fAnzPackungen! * 10) = CLng(fAnzPackungen! * 10)) Then
'            h$ = Format(fAnzPackungen!, "0.0")
'        Else
'            h$ = Format(fAnzPackungen!, "0.00")
'        End If
'    End If
'    .TextMatrix(row%, 1) = h$
'
'    h$ = ""
'    'neu in 3.0.71
'    If (fAnzPackungen! <> 0) Then
'        dEK = Abs(dWert / fAnzPackungen)
'    End If
'    If (dEK <> 0) Then
'        h$ = Format(dEK, "0.00")
'    End If
'    .TextMatrix(row%, 5) = h$
'
'    h$ = ""
'    If (dWert# <> 0) Then
'        h$ = Format(dWert#, "0.00")
'    End If
'    .TextMatrix(row%, 7) = h$
'
'    h$ = ""
'    If (dWertInkl# <> 0) Then
'        h$ = Format(dWertInkl#, "0.00")
'    End If
'    .TextMatrix(row%, 8) = h$
'End With
    
Call DefErrPop
End Sub

Sub ZeigeBelegZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBelegZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, row%
Dim fAnzPackungen!
Dim dWert#, dWertInkl#
Dim h$, h2$, pzn$

'With flxarbeit(0)
'    If (FaktSQL) Then
'        pzn$ = PznString(FArtAdoRec!pzn)
'        fAnzPackungen! = FArtAdoRec!Anzahl
'        dWert# = FArtAdoRec!Wert
'        dWertInkl# = FArtAdoRec!WertInkl
'
'        row% = 0
'        For i% = .FixedRows To (.Rows - 1)
'            If (pzn$ = .TextMatrix(i%, 0)) Then
'                row% = i%
'                If (Trim(pzn$) = "") Or (Val(pzn$) = 9999999) Then
'                    If (FArtAdoRec!text <> .TextMatrix(i%, 2)) Then
'                        row% = 0
'                    ElseIf (FArtAdoRec!menge <> .TextMatrix(i%, 3)) Then
'                        row% = 0
'                    ElseIf (FArtAdoRec!einheit <> .TextMatrix(i%, 4)) Then
'                        row% = 0
'                    End If
'                End If
'                If (row%) Then
'                    Exit For
'                End If
'            End If
'        Next i%
'
'
'        If (row% = 0) Then
'            .AddItem ""
'            row% = .Rows - 1
'
'            .TextMatrix(row%, 0) = pzn$
'
'            .TextMatrix(row%, 2) = FArtAdoRec!text
'            .TextMatrix(row%, 3) = FArtAdoRec!menge
'            .TextMatrix(row%, 4) = FArtAdoRec!einheit
'
'            h$ = .TextMatrix(row%, 0)
'            For i% = 2 To 4
'                h$ = h$ + .TextMatrix(row%, i%)
'            Next i%
'            .TextMatrix(row%, 12) = h$
'        Else
'            fAnzPackungen! = xVal(.TextMatrix(row%, 1)) + fAnzPackungen!
'            dWert# = xVal(.TextMatrix(row%, 7)) + FArtAdoRec!Wert
'            dWertInkl# = xVal(.TextMatrix(row%, 8)) + FArtAdoRec!WertInkl
'        End If
'    Else
'        pzn$ = PznString(FArtRec!pzn)
'        fAnzPackungen! = FArtRec!Anzahl
'        dWert# = FArtRec!Wert
'        dWertInkl# = FArtRec!WertInkl
'
'        row% = 0
'        For i% = .FixedRows To (.Rows - 1)
'            If (pzn$ = .TextMatrix(i%, 0)) Then
'                row% = i%
'                If (Trim(pzn$) = "") Or (Val(pzn$) = 9999999) Then
'                    If (FArtRec!text <> .TextMatrix(i%, 2)) Then
'                        row% = 0
'                    ElseIf (FArtRec!menge <> .TextMatrix(i%, 3)) Then
'                        row% = 0
'                    ElseIf (FArtRec!einheit <> .TextMatrix(i%, 4)) Then
'                        row% = 0
'                    End If
'                End If
'                If (row%) Then
'                    Exit For
'                End If
'            End If
'        Next i%
'
'
'        If (row% = 0) Then
'            .AddItem ""
'            row% = .Rows - 1
'
'            .TextMatrix(row%, 0) = pzn$
'
'            .TextMatrix(row%, 2) = FArtRec!text
'            .TextMatrix(row%, 3) = FArtRec!menge
'            .TextMatrix(row%, 4) = FArtRec!einheit
'
'            h$ = .TextMatrix(row%, 0)
'            For i% = 2 To 4
'                h$ = h$ + .TextMatrix(row%, i%)
'            Next i%
'            .TextMatrix(row%, 12) = h$
'        Else
'            fAnzPackungen! = xVal(.TextMatrix(row%, 1)) + fAnzPackungen!
'            dWert# = xVal(.TextMatrix(row%, 7)) + FArtRec!Wert
'            dWertInkl# = xVal(.TextMatrix(row%, 8)) + FArtRec!WertInkl
'        End If
'    End If
'
'    h$ = ""
'    If (fAnzPackungen! <> 0) Then
'        If (fAnzPackungen! = CLng(fAnzPackungen!)) Then
'            h$ = Format(fAnzPackungen!, "0")
'        ElseIf (FNX(fAnzPackungen! * 10) = CLng(fAnzPackungen! * 10)) Then
'            h$ = Format(fAnzPackungen!, "0.0")
'        Else
'            h$ = Format(fAnzPackungen!, "0.00")
'        End If
'    End If
'    .TextMatrix(row%, 1) = h$
'
'    h$ = ""
'    If (dWert# <> 0) Then
'        h$ = Format(dWert#, "0.00")
'    End If
'    .TextMatrix(row%, 7) = h$
'
'    h$ = ""
'    If (dWertInkl# <> 0) Then
'        h$ = Format(dWertInkl#, "0.00")
'    End If
'    .TextMatrix(row%, 8) = h$
'End With
    
Call DefErrPop
End Sub

Sub AuswahlSortieren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlSortieren")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim dProz#, dSumProz#
Dim h$

With flxarbeit(0)
    .Redraw = False
    
    .row = .FixedRows
    .col = 8
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 4   'Zahlen absteigend
    
    .Redraw = True
    .TopRow = .FixedRows
    .row = .FixedRows
    .col = 0
    .ColSel = .Cols - 1
'    If (.Visible) Then
'        .SetFocus
'    End If

'    Call FormKurzInfo
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
'    erg% = 0
'    If (pzn$ <> "") Then
'        erg% = artstat.StatistikRechnen(pzn$)
'    End If
'    If (erg%) Then
'        AltJahr& = -1
'        j% = 0
'        For i% = 0 To 12
'            Termin& = artstat.Anfang - i% - 1
'            Jahr& = (Termin& - 1) \ 12
'            Monat% = Termin& - Jahr& * 12
'            If (AltJahr& <> Jahr&) Then
'                .TextMatrix(0, j%) = Str$(Jahr&)
'                iWert! = artstat.JahresWert(Jahr&)
'                If (iWert! <> 0!) Then
'                    .TextMatrix(1, j%) = Str$(iWert!)
'                Else
'                    .TextMatrix(1, j%) = ""
'                End If
'                .row = 0
'                .col = j%
'                .CellFontBold = True
'                .CellFontSize = .Font.Size
'                .row = 1
'                .CellBackColor = .BackColorFixed
'                .CellFontBold = True
'                .CellFontSize = .Font.Size
'                j% = j% + 1
'                AltJahr& = Jahr&
'            End If
'
'            .TextMatrix(0, j%) = para.MonatKurz(Monat%)
'
'            iWert! = artstat.MonatsWert(i% + 1)
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
'
'    .Redraw = True
'End With

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


Private Sub picBack_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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
Static OrgX!, OrgY!

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (para.Newline) Then
    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
        OrgX = x
        OrgY = y
        
        Call opToolbar.ShowToolbar
    End If
End If

Call DefErrPop
End Sub

Private Sub flxarbeit_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxarbeit_MouseMove")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Static OrgX!, OrgY!

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If (para.Newline) Then
    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
        OrgX = x
        OrgY = y
        
        Call opToolbar.ShowToolbar
    End If
End If

Call DefErrPop
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

If (para.Newline) Then
'    Call opToolbar.ShowToolbar
    Call opToolbar.Click(x)
Else
    picToolbar.Drag (vbBeginDrag)
    opToolbar.DragX = x
    opToolbar.DragY = y
End If

Call DefErrPop
End Sub

Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picToolbar_MouseMove")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, cmdToolbarSize%, xx%, loch%, IconWidth%, index%
Dim h$
Static OrgX!, OrgY!

If (para.Newline) Then
    If (Abs(x - OrgX) > 15) Or (Abs(y - OrgY) > 15) Then
        OrgX = x
        OrgY = y
        
'        Call opToolbar.ShowToolbar
        Call opToolbar.MouseMove(x)
    End If
End If

Call DefErrPop
End Sub

Sub ZerlegeBezügeKundenNummern(sKundenNummern$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZerlegeBezügeKundenNummern")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim kk%, ind%
Dim kMax&, iKuNr&
Dim nummer1!, nummer2!
Dim x$, k$

'kMax = 0
'SQLStr$ = "SELECT max(KundenNr) as MaxKuNr FROM Kunden"
'#If (KUNDEN_SQL = -1) Then
'    On Error Resume Next
'    KundenRec.Close
'    Err.Clear
'    On Error GoTo DefErr
'    KundenRec.open SQLStr, KundenAdoDB.ActiveConn
'#Else
'    Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
'#End If
'If (KundenRec.EOF = False) Then
'    kMax = CheckNullLong(KundenRec!MaxKuNr)
'End If
'
'For kk% = 1 To 20
'    BezügeKnrVon(kk%) = 0
'    BezügeKnrBis(kk%) = 0
'Next kk%
'BezügeSammelKnr = ""
'
'If (sKundenNummern$ <> "") Then
'    x$ = sKundenNummern$
'    kk% = 1
'    While (kk% <= Len(x$))
'        ind% = InStr("0123456789-,", Mid$(x$, kk%, 1))
'        If (ind% = 0) Then
'            x$ = Left$(x$, kk% - 1) + Mid$(x$, kk% + 1)
'        Else
'            kk% = kk% + 1
'        End If
'    Wend
'
'    If (Right$(x$, 1) <> ",") Then
'        x$ = x$ + ","
'    End If
'
'    kk% = 0
'    Do
'        ind% = InStr(x$, ",")
'        If (ind% > 0) Then
'            If (kk% < 20) Then
'                k$ = Left$(x$, ind% - 1)
'                x$ = Mid$(x$, ind% + 1)
'                ind% = InStr(k$, "-")
'                If (ind% > 0) Then
'                    nummer1! = Val(Left$(k$, ind% - 1))
'                    nummer2! = Val(Mid$(k$, ind% + 1))
'                Else
'                    nummer1! = Val(k$)
'                    nummer2! = nummer1!
'                End If
'                If (nummer1! <= kMax) And (nummer2! <= kMax) Then
'                    kk% = kk% + 1
'                    BezügeKnrVon(kk%) = nummer1!
'                    BezügeKnrBis(kk%) = nummer2!
'                    For iKuNr = nummer1 To nummer2
'                        SQLStr$ = "SELECT * FROM Kunden WHERE SammelKundenNr=" + Str$(iKuNr)
'                        #If (KUNDEN_SQL = -1) Then
'                            On Error Resume Next
'                            KundenRec.Close
'                            Err.Clear
'                            On Error GoTo DefErr
'                            KundenRec.open SQLStr, KundenAdoDB.ActiveConn
'                        #Else
'                            Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
'                        #End If
'                        Do
'                            If (KundenRec.EOF) Then
'                                Exit Do
'                            End If
'
'                            If (InStr("," + BezügeSammelKnr, "," + CStr(KundenRec!KundenNr) + ",") <= 0) Then
'                                BezügeSammelKnr = BezügeSammelKnr + CStr(KundenRec!KundenNr) + ","
'                            End If
'
'                            KundenRec.MoveNext
'                        Loop
'                    Next
'                End If
'            Else
'                Exit Do
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'
'    If (BezügeSammelKnr <> "") Then
'        BezügeSammelKnr = "," + BezügeSammelKnr
'    End If
'End If

Call DefErrPop
End Sub

Function CheckBezügeKundenNummer%(iKuNr)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBezügeKundenNummer%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim kk%, ret%
Dim iKuNr2&
Dim sAuswahlKz$

'ret% = True
'If (ParamKunden$ <> "") Then
'    ret% = 0
'    For kk% = 1 To 20
'        If (iKuNr >= BezügeKnrVon(kk%)) And (iKuNr <= BezügeKnrBis(kk%)) Then
'            ret% = True
'            Exit For
'        End If
'    Next kk%
'    If (ret% = 0) And (ParamSammelbeleg%) Then
'        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(iKuNr)
'        #If (KUNDEN_SQL = -1) Then
'            On Error Resume Next
'            KundenRec.Close
'            Err.Clear
'            On Error GoTo DefErr
'            KundenRec.open SQLStr, KundenAdoDB.ActiveConn
'        #Else
'            Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
'        #End If
'        If (KundenRec.EOF = False) Then
'            If (KundenRec!SammelKundenNr > 0) Then
'                iKuNr2 = KundenRec!SammelKundenNr
'                For kk% = 1 To 20
'                    If (iKuNr2 >= BezügeKnrVon(kk%)) And (iKuNr2 <= BezügeKnrBis(kk%)) Then
'                        ret% = True
'                        Exit For
'                    End If
'                Next kk%
'            End If
'    '        If (BelegKuNr% = kund.SammelKnr) Then
'    '            ret% = True
'    '        End If
'        End If
'    End If
'End If
'
'If (ret%) And (ParamAuswahlKz$ <> "") Then
'    SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(iKuNr)
'    #If (KUNDEN_SQL = -1) Then
'        On Error Resume Next
'        KundenRec.Close
'        Err.Clear
'        On Error GoTo DefErr
'        KundenRec.open SQLStr, KundenAdoDB.ActiveConn
'    #Else
'        Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
'    #End If
'    If (KundenRec.EOF) Then
'        ret% = 0
'    Else
'        sAuswahlKz$ = UCase$(CheckNullStr(KundenRec!AuswahlKz))
''        For kk% = 1 To Len(ParamAuswahlKz$)
''            If (InStr(sAuswahlKz$, Mid$(ParamAuswahlKz$, kk%, 1)) = 0) Then
''                ret% = 0
''                Exit For
''            End If
''        Next kk%
'        ret% = (InStr(sAuswahlKz$, ParamAuswahlKz$) > 0)
'    End If
'End If

CheckBezügeKundenNummer% = ret%

Call DefErrPop
End Function

Sub AuswahlSetzen(Optional iHakerl% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlSetzen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim h$

If (iHakerl%) Then
    h$ = Chr$(214)
Else
    h$ = ""
End If

With flxarbeit(0)
    .Redraw = False
    
'    Call HighlightZeile(True)
    
    For i% = .FixedRows To (.Rows - 1)
        .row = i%
        .TextMatrix(i%, 0) = h$
    Next i%
    
    .Redraw = True
    .TopRow = .FixedRows
    .row = .FixedRows
    .col = 0
    .ColSel = .Cols - 1
    If (.Visible) Then
        .SetFocus
    End If
    
    Call FormKurzInfo
End With

Call DefErrPop
End Sub

Sub ToggleAuswahlZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToggleAuswahlZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim h$

KeinRowColChange% = True
With flxarbeit(0)
    h$ = .TextMatrix(.row, 0)
    If (h$ <> "") Then
        h$ = ""
    Else
        h$ = Chr$(214)
    End If
    .TextMatrix(.row, 0) = h$
    
    .col = 0
    .ColSel = .Cols - 1
End With
KeinRowColChange% = False

'Call ActProgram.SumArtikelListe

Call DefErrPop
End Sub

'Sub NaechsteAuswahlZeile()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("NaechsteAuswahlZeile")
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
'Dim iArbeitAnzzeilen%
'
'With flxarbeit(0)
'    If (.row < .Rows - 1) Then
'        .row = .row + 1
'        iArbeitAnzzeilen% = (.Height - 90) \ .RowHeight(0)
'        If (.TopRow + iArbeitAnzzeilen - 1 <= .row) Then
'            .TopRow = .row
'        End If
'    End If
'    .col = 0
'    .ColSel = .Cols - 1
'End With
'
'Call DefErrPop
'End Sub

Private Sub cboAuswertung_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAuswertung_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim WoTag%, iQuart%
Dim h$
Dim iDat As Date

If Not (cboAuswertung(index).Visible) Then
    Call DefErrPop: Exit Sub
End If

txtAuswertung(2 * index).Visible = False
txtAuswertung(2 * index + 1).Visible = False
txtAuswertung(2 * index).text = ""
txtAuswertung(2 * index + 1).text = ""
h = UCase(Trim(cboAuswertung(index).text))
If (h <> "") Then
'    If (index = 2) Or (index = 4) Or (index = 5) Or (index = 6) Or (index = 7) Or (index = 10) Or (index = 16) Then
        With txtAuswertung(2 * index)
            .Visible = True
            .BackColor = vbWhite
            With .Container
                .ForeColor = RGB(180, 180, 180) ' vbWhite
                .FillStyle = vbSolid
                .FillColor = vbWhite

                RoundRect .hdc, (txtAuswertung(2 * index).Left - 60) / Screen.TwipsPerPixelX, (txtAuswertung(2 * index).Top - 30) / Screen.TwipsPerPixelY, (txtAuswertung(2 * index).Left + txtAuswertung(2 * index).Width + 60) / Screen.TwipsPerPixelX, (txtAuswertung(2 * index).Top + txtAuswertung(2 * index).Height + 15) / Screen.TwipsPerPixelY, 10, 10
            End With
        
'            txtAuswertung(5).MaxLength = 6
'            If (index = 16) Then
'                txtAuswertung(2 * index + 1).MaxLength = 6
'            End If
            
            If (index = 0) And (h = "=") Then
                txtAuswertung(2 * index + 1).Visible = True
                txtAuswertung(2 * index + 1).MaxLength = 4
            ElseIf (h = UCase("Zwischen")) Then
                txtAuswertung(2 * index + 1).Visible = True
            ElseIf (h = UCase("Heute")) Then
                cboAuswertung(index).ListIndex = 3
                .text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("Gestern")) Then
                cboAuswertung(index).ListIndex = 3
                .text = Format(Now - 1, "DDMMYY")
            ElseIf (h = UCase("diese Woche")) Then
                txtAuswertung(2 * index + 1).Visible = True
                WoTag = Weekday(Now, vbMonday)
                .text = Format(Now - WoTag + 1, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now - WoTag + 1 + 6, "DDMMYY")
            ElseIf (h = UCase("letzte Woche")) Then
                txtAuswertung(2 * index + 1).Visible = True
                WoTag = Weekday(Now, vbMonday)
                .text = Format(Now - WoTag + 1 - 7, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now - WoTag + 1 + 6 - 7, "DDMMYY")
            ElseIf (h = UCase("letzte 7 Tage")) Then
                txtAuswertung(2 * index + 1).Visible = True
                WoTag = Weekday(Now, vbMonday)
                .text = Format(Now - 6, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("dieses Monat")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = Now
                .text = "01" + Format(iDat, "MMYY")
                txtAuswertung(2 * index + 1).text = Format(DateAdd("d", -1, "01." + Format(DateAdd("m", 1, iDat), "MM.YYYY")), "DDMMYY")
            ElseIf (h = UCase("letztes Monat")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = DateAdd("m", -1, Now)
                .text = "01" + Format(iDat, "MMYY")
                txtAuswertung(2 * index + 1).text = Format(DateAdd("d", -1, "01." + Format(DateAdd("m", 1, iDat), "MM.YYYY")), "DDMMYY")
            ElseIf (h = UCase("letzte 4 Wochen")) Then
                txtAuswertung(2 * index + 1).Visible = True
                WoTag = Weekday(Now, vbMonday)
                .text = Format(Now - WoTag + 1 - 28, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("letzte 30 Tage")) Then
                txtAuswertung(2 * index + 1).Visible = True
                .text = Format(Now - 30, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("dieses Quartal")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = Now
                iQuart = Val(Format(iDat, "q"))
                iDat = "01." + Format((iQuart - 1) * 3 + 1, "00") + "." + Format(iDat, "yyyy")
                .text = Format(iDat, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(DateAdd("d", -1, "01." + Format(DateAdd("m", 3, iDat), "MM.YYYY")), "DDMMYY")
            ElseIf (h = UCase("letztes Quartal")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = DateAdd("m", -3, Now)
                iQuart = Val(Format(iDat, "q"))
                iDat = "01." + Format((iQuart - 1) * 3 + 1, "00") + "." + Format(iDat, "yyyy")
                .text = Format(iDat, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(DateAdd("d", -1, "01." + Format(DateAdd("m", 3, iDat), "MM.YYYY")), "DDMMYY")
            ElseIf (h = UCase("dieses Jahr")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = Now
                .text = "0101" + Format(iDat, "YY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("letztes Jahr")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = DateAdd("yyyy", -1, Now)
                .text = "0101" + Format(iDat, "YY")
                txtAuswertung(2 * index + 1).text = "3112" + Format(iDat, "YY")
            ElseIf (h = UCase("letzte 12 Monate")) Then
                txtAuswertung(2 * index + 1).Visible = True
                iDat = DateAdd("yyyy", -1, Now)
                .text = Format(iDat, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("letzte 52 Wochen")) Then
                txtAuswertung(2 * index + 1).Visible = True
                WoTag = Weekday(Now, vbMonday)
                .text = Format(Now - WoTag + 1 - 52 * 7, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            ElseIf (h = UCase("letzte 365 Tage")) Then
                txtAuswertung(2 * index + 1).Visible = True
                .text = Format(Now - 365, "DDMMYY")
                txtAuswertung(2 * index + 1).text = Format(Now, "DDMMYY")
            End If
            
            .SetFocus
        End With
        
        With txtAuswertung(2 * index + 1)
            If (.Visible) Then
                .BackColor = vbWhite
                With .Container
                    .ForeColor = RGB(180, 180, 180) ' vbWhite
                    .FillStyle = vbSolid
                    .FillColor = vbWhite
    
                    RoundRect .hdc, (txtAuswertung(2 * index + 1).Left - 60) / Screen.TwipsPerPixelX, (txtAuswertung(2 * index + 1).Top - 30) / Screen.TwipsPerPixelY, (txtAuswertung(2 * index + 1).Left + txtAuswertung(2 * index + 1).Width + 60) / Screen.TwipsPerPixelX, (txtAuswertung(2 * index).Top + txtAuswertung(2 * index).Height + 15) / Screen.TwipsPerPixelY, 10, 10
                End With
            End If
        End With
'    End If
End If

Call DefErrPop
End Sub

Sub Drucke_eRezepte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Drucke_eRezepte")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, ind%, ze%, y%
Dim gesBreite&, x&
Dim h$, h2$, header2$, sDruckerOrg$

Dim hPrinter As Printer
sDruckerOrg = Printer.DeviceName
For Each hPrinter In Printers
    If (UCase(StandardDrucker) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next


Printer.Orientation = vbPRORLandscape   ' vbPRORPortrait

Call StartAnimation(Me, "Ausdruck wird erstellt ...")

header2$ = cboAuswahl(0).text

'''''''''''''''
AnzDruckSpalten% = 28
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .Titel = "Gescannt"
    .TypStr = String$(11, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(1)
    .Titel = "TaskId"
    .TypStr = String$(15, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(2)
    .Titel = "RezeptNr"
    .TypStr = String$(13, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(3)
    .Titel = "VO#"
    .TypStr = String$(4, "X")
    .Ausrichtung = "Z"
End With
With DruckSpalte(4)
    .Titel = "ABG#"
    .TypStr = String$(4, "X")
    .Ausrichtung = "Z"
End With
With DruckSpalte(5)
    .Titel = "PZN"
    .TypStr = String$(8, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(6)
    .Titel = "A R T I K E L"
    .TypStr = String$(30, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(7)
    .Titel = ""
    .TypStr = String$(6, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(8)
    .Titel = ""
    .TypStr = String$(5, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(9)
    .Titel = "Dar"
    .TypStr = String$(4, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(10)
    .Titel = "N"
    .TypStr = String$(3, "X")
    .Ausrichtung = "Z"
End With
With DruckSpalte(11)
    .Titel = "ChargenNr"
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(12)
    .Titel = "Doierung"
    .TypStr = String$(10, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(13)
    .Titel = "AI#"
    .TypStr = String$(3, "X")
    .Ausrichtung = "Z"
    .Attrib = 2
End With
With DruckSpalte(14)
    .Titel = "KK"
    .TypStr = String$(8, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(15)
    .Titel = "KK-Name"
    .TypStr = String$(15, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(16)
    .Titel = "ZuzFrei"
    .TypStr = String$(4, "X")
    .Ausrichtung = "Z"
    .Attrib = 2
End With
With DruckSpalte(17)
    .Titel = "Versicherter"
    .TypStr = String$(9, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(18)
    .Titel = "Patient"
    .TypStr = String$(17, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(19)
    .Titel = "Impf"
    .TypStr = String$(3, "X")
    .Ausrichtung = "Z"
    .Attrib = 2
End With
With DruckSpalte(20)
    .Titel = "BVG"
    .TypStr = String$(3, "X")
    .Ausrichtung = "Z"
    .Attrib = 2
End With
With DruckSpalte(21)
    .Titel = "Noct"
    .TypStr = String$(3, "X")
    .Ausrichtung = "Z"
    .Attrib = 2
End With
With DruckSpalte(22)
    .Titel = "AbrechBis"
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(23)
    .Titel = "EinreichBis"
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(24)
    .Titel = "GesBrutto"
    .TypStr = String$(8, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(25)
    .Titel = "Zuzahl"
    .TypStr = String$(6, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(26)
    .Titel = "ALBVVG"
    .TypStr = String$(6, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(27)
    .Titel = "AnlageAm"
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With

Call InitDruckZeile(True)

DruckSeite% = 0
ze% = 0
Call eRezepte_DruckKopf

With flxarbeit(0)
    For i% = 1 To (.Rows - 1)
        h = ""
        h = h + .TextMatrix(i, 1) + vbTab
        h = h + .TextMatrix(i, 2) + vbTab
        h = h + .TextMatrix(i, 3) + vbTab
        h = h + .TextMatrix(i, 5) + vbTab
        h = h + .TextMatrix(i, 6) + vbTab
        h = h + .TextMatrix(i, 7) + vbTab
        h = h + .TextMatrix(i, 8) + vbTab
        h = h + .TextMatrix(i, 9) + vbTab
        h = h + .TextMatrix(i, 10) + vbTab
        h = h + .TextMatrix(i, 11) + vbTab
        h = h + .TextMatrix(i, 12) + vbTab
        h = h + .TextMatrix(i, 13) + vbTab
        h = h + .TextMatrix(i, 14) + vbTab
        h = h + .TextMatrix(i, 15) + vbTab
        h = h + .TextMatrix(i, 16) + vbTab
        h = h + .TextMatrix(i, 17) + vbTab
        h = h + .TextMatrix(i, 19) + vbTab
        h = h + .TextMatrix(i, 20) + vbTab
        h = h + .TextMatrix(i, 21) + vbTab
        h = h + .TextMatrix(i, 30) + vbTab
        h = h + .TextMatrix(i, 31) + vbTab
        h = h + .TextMatrix(i, 32) + vbTab
        h = h + .TextMatrix(i, 33) + vbTab
        h = h + .TextMatrix(i, 34) + vbTab
        h = h + .TextMatrix(i, 41) + vbTab
        h = h + .TextMatrix(i, 42) + vbTab
        h = h + .TextMatrix(i, 43) + vbTab
        h = h + .TextMatrix(i, 38) + vbTab
        
        Call DruckZeile(h$)
        
        ze% = ze% + 1
        If (ze% Mod 3 = 0) Then
            Printer.CurrentY = Printer.CurrentY + 60
        End If
                
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call eRezepte_DruckKopf
            ze% = 0
        End If
    Next i%
End With
    
With flxSummen
    If (.Visible) Then
        y% = Printer.CurrentY + 30
        gesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
        Printer.Line (DruckSpalte(23).StartX, y%)-(gesBreite&, y%)
        Printer.CurrentY = y% + 60
    
        h$ = ""
'        For j% = 0 To (AnzDruckSpalten% - 1)
'            h$ = h$ + .TextMatrix(i%, j%) + vbTab
'        Next j%
        For j% = 0 To 22
            h$ = h$ + vbTab
        Next j%
        h$ = h$ + .TextMatrix(0, 2) + " eRez" + vbTab
        h$ = h$ + .TextMatrix(0, 4) + vbTab
        h$ = h$ + .TextMatrix(0, 6) + vbTab
        h$ = h$ + .TextMatrix(0, 8) + vbTab
        h$ = h$ + vbTab
        Printer.FontBold = True
        Call DruckZeile(h$)
        Printer.FontBold = False
    End If
End With
    
'y% = Printer.CurrentY + 90
'GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
'Printer.Line (DruckSpalte(0).StartX, y%)-(GesBreite&, y%)
'Printer.CurrentY = y% + 90
'With flxSummen
'    h$ = vbTab + .TextMatrix(0, 0)
'    For j% = 0 To 4
'        h$ = h$ + vbTab
'    Next j%
'    If (ParamKundenSort) Then
'        For j% = 1 To 1
'            h$ = h$ + vbTab + .TextMatrix(0, j%)
'        Next j%
'        For j% = 1 To 4
'            h$ = h$ + vbTab
'        Next j%
'    Else
'        For j% = 1 To 2
'            h$ = h$ + vbTab + .TextMatrix(0, j%)
'        Next j%
'        For j% = 1 To 3
'            h$ = h$ + vbTab
'        Next j%
'    End If
'End With
'Call DruckZeile(h$)

Call DruckFuss(False)
Printer.EndDoc
                
For Each hPrinter In Printers
    If (UCase(sDruckerOrg) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next
              
Call StopAnimation(Me)

Call DefErrPop
End Sub

Sub eRezepte_DruckKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("eRezepte_DruckKopf")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, x%, y%
Dim gesBreite&
Dim header$, KopfZeile$, Typ$, h$

'KopfZeile$ = "Import-Kontrolle"
'header$ = "Import-Kontrolle" + " " + cmbDatum.List(cmbDatum.ListIndex)
KopfZeile$ = Me.Caption
header$ = KopfZeile$ + ": " + Trim(Left(cboAuswahl(0).text, 50)) + ", " + cboAuswertung(0).text
With txtAuswertung(0)
    If (.Visible) Then
        header = header + " " + .text
    End If
End With
With txtAuswertung(1)
    If (.Visible) Then
        header = header + " - " + .text
    End If
End With
Call DruckKopf(header$, Typ$, KopfZeile$, 0)
Printer.CurrentY = Printer.CurrentY - 3 * Printer.TextHeight("A")
    
For i% = 0 To (AnzDruckSpalten% - 1)
    h$ = RTrim(DruckSpalte(i%).Titel)
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h$)
    End If
    Printer.CurrentX = x%
    Printer.Print h$;
Next i%

Printer.Print " "

y% = Printer.CurrentY
gesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, y%)-(gesBreite&, y%)

y% = Printer.CurrentY
Printer.CurrentY = y% + 30

Call DefErrPop

End Sub


Sub Export_eRezepte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Export_eRezepte")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim i%, j%, ind%, erg%, EXPORT_ID%
Dim sDatei$, h$, txt$, ParameterCSV$, sName$

sName = "eRezepte.csv"
Do
    sName = Trim(MyInputBox("Name der Excel-CSV-Datei:", Me.Caption, sName))
    If (sName = "") Then
        Call DefErrPop: Exit Sub
    End If

    ParameterCSV = sName
    Exit Do
Loop
If (Right(UCase(ParameterCSV), 4) <> ".CSV") Then
    ParameterCSV = ParameterCSV + ".csv"
End If


h = CurDir + "\eRezepte_csv": erg% = wpara.CreateDirectory(h)
sDatei = h$ + "\" + ParameterCSV '+ ".csv"
EXPORT_ID% = FileOpen(sDatei, "O")

With flxarbeit(0)
    .Redraw = False
    For i = 0 To (.Rows - 1)
        txt = ""
        For j = 1 To (.Cols - 1)
            If (.ColWidth(j) > 0) Or (j = 2) Then   'auch die Task-Nr
                h = .TextMatrix(i, j)
                If (i > 0) Then
                    .row = i
                    .col = j
                    If (.CellFontName = "Symbol") Then
                        h = IIf(.TextMatrix(i, j) <> "", "1", "")
                    End If
                End If
                If (j = 3) Then 'Rezept-Nr
                    txt = txt + "=" + Chr(34) + h + Chr(34) + ";"
                Else
                    txt = txt + h + ";"
                End If
            End If
        Next
        Print #EXPORT_ID%, txt$
    Next i
    
    If (flxSummen.Visible) Then
        Print #EXPORT_ID%,
        
        txt = ""
        For j = 1 To (.Cols - 8)
            If (.ColWidth(j) > 0) Or (j = 2) Then   'auch die Task-Nr
                txt = txt + ";"
            End If
        Next
        
        txt = txt + flxSummen.TextMatrix(0, 2) + " eRez" + ";"
        txt = txt + flxSummen.TextMatrix(0, 4) + ";"
        txt = txt + flxSummen.TextMatrix(0, 6) + ";"
        txt = txt + flxSummen.TextMatrix(0, 8) + ";"
        Print #EXPORT_ID%, txt$
    End If
    
    .Redraw = True
End With

    
Call MessageBox("Datei " + sDatei + " wurde erzeugt!", vbInformation)
Close #EXPORT_ID%

Call DefErrPop
End Sub

Private Sub flxArbeit_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeit_MouseUp")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, col%, row%, iSortModus%, OrgSortCol%, SortNeu%
Dim KeyCol%
Dim dKey#
Dim h$, Key$
Dim iDat As Date

If (ProgrammModus = 1) Then
    Call DefErrPop: Exit Sub
End If

OrgSortCol% = SortCol%
SortNeu% = True
With flxarbeit(0)
    KeyCol% = .Cols - 2
    
    row% = .row
    
    If (.Rows <= 2) Then
        Call DefErrPop: Exit Sub
    End If
    
    If (y <= .RowHeight(0)) Then
        col% = .Cols - 1
        For i% = 1 To (.Cols - 1)
            If (x < .ColPos(i%)) Then
                col% = i% - 1
                Exit For
            End If
        Next i%
        
        SortCol% = col%
        
        .Redraw = False
        
        If ((Shift And 2) = 0) Then
            .FillStyle = flexFillRepeat
            .row = 1
            .col = KeyCol%
            .RowSel = .Rows - 1
            .ColSel = .col
            .text = ""
        
            .row = 0
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
'            .CellFontBold = False
            .CellBackColor = .BackColor
            .FillStyle = flexFillSingle
        
            If (SortCol% <> OrgSortCol%) Then
                SortModus% = 0
            Else
                SortModus% = (SortModus% + 1) Mod 2
            End If
            
            AnzSortCols% = 1
            SortCols%(0) = SortCol%
        ElseIf (SortCol% = OrgSortCol%) Then
            SortModus% = (SortModus% + 1) Mod 2
            SortNeu% = 0
        Else
            SortCols%(AnzSortCols%) = SortCol%
            AnzSortCols% = AnzSortCols% + 1
        End If
        
        If (SortNeu%) Then
            For i% = 1 To (.Rows - 1)
                iSortModus% = Mid(SortTypen, col% + 1, 1)
                h$ = .TextMatrix(i%, col%)
                
                Select Case iSortModus%
                
                    Case 0
                        Key$ = h$
                    Case 1
'                        dKey# = xVal(h$)
'                        If (dKey# >= 0) Then
'                            key$ = Format(dKey#, "+000000.00")
'                        Else
'                            key$ = Format(dKey#, "000000.00")
'                        End If
                        dKey# = xVal(h$) + 100000000
'                        If (dKey# >= 0) Then
'                            key$ = Format(dKey#, "+000000.00")
'                        Else
                            Key$ = Format(dKey#, "000000000.00")
'                        End If
                    Case 2
                        If (h$ = "") Then
                            iDat = "01.01.1900"
                        Else
'                            iDat = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "." + "20" + Right$(h$, 2)
                            iDat = h$
                        End If
                        Key$ = Format(iDat, "YYYYMMDDHHmm")
                    Case 3
                        If (h$ = Chr$(214)) Then
                            Key$ = "0"
                        Else
                            Key$ = "1"
                        End If
                    
                End Select
                
                .TextMatrix(i%, KeyCol%) = .TextMatrix(i%, KeyCol%) + Key$
            Next i%
        End If
        
        .row = 1
        .col = KeyCol%
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 5 + SortModus%
        .TopRow = 1
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
        
        .FillStyle = flexFillRepeat
        .row = 0
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellBackColor = .BackColorFixed
        .row = 0
        .col = SortCol%
        .RowSel = .row
        .ColSel = .col
'        .CellFontBold = True
        .CellBackColor = RGB(255, 170, 120) ' vbMagenta
        .FillStyle = flexFillSingle
        
        .row = row%
        .col = 0
        .ColSel = .Cols - 1
        
        .Redraw = True
    End If
End With

Call DefErrPop
End Sub

Private Function OpStatusStr(OpStatus As Integer) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpStatusStr")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim sRet$

sRet = "(" + CStr(OpStatus) + ")"
Select Case OpStatus
    Case 0, 1
        sRet = "Zugewiesen"
    Case 2
        sRet = "Eingescannt"
    Case 3
        sRet = "Dispensiert"
    Case 4
        sRet = "Eingereicht"
    Case 5
        sRet = "Fehler"
    Case 6
        sRet = "Abrechenbar"
    Case 7
        sRet = "Ohne weitere Bearbeitung"
End Select

OpStatusStr = sRet

Call DefErrPop
End Function

Private Sub CreateQR()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateQR")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim sTaskId$, sAccessCode$, sQR$

With flxarbeit(0)
    row% = .row
    sTaskId = .TextMatrix(row, 2)
    sAccessCode = .TextMatrix(row, 25)
    If (sTaskId <> "") And (sAccessCode <> "") Then
        sQR = sQR + IIf(sQR = "", "", ",") + Chr(34) + "Task/" + sTaskId + "/$accept?ac=" + sAccessCode + Chr(34)
        sQR = "{" + Chr(34) + "urls" + Chr(34) + ":[" + sQR + "]}"
        Clipboard.SetText sQR
        Call MsgBox("Der QR-Code" + vbCrLf + vbCrLf + sQR + vbCrLf + vbCrLf + "wurde in die Zwischenablage kopiert!", vbInformation)
    End If
End With

Call DefErrPop
End Sub

