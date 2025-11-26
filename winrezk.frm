VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmAction 
   Caption         =   "Rezeptkontrolle"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "winrezk.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   8580
   ScaleWidth      =   16140
   Begin SHDocVwCtl.WebBrowser wbAbgabedaten 
      Height          =   2535
      Left            =   9240
      TabIndex        =   39
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
   Begin SHDocVwCtl.WebBrowser wbVerordnung 
      Height          =   2535
      Left            =   3960
      TabIndex        =   40
      Top             =   1560
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
   Begin VB.ListBox lstMarsRezepte 
      Height          =   300
      Left            =   12240
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrF6Sperre 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   13680
      Top             =   240
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
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   11880
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picGDI 
      Height          =   735
      Index           =   0
      Left            =   5040
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picEan 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1335
      Left            =   -240
      ScaleHeight     =   1335
      ScaleWidth      =   3495
      TabIndex        =   29
      Top             =   3960
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picTaxierungsDruck 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   315
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdDatei 
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox picSave 
      Height          =   615
      Left            =   4080
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   16
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
      TabIndex        =   3
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
         TabIndex        =   4
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
      Left            =   3000
      ScaleHeight     =   2370
      ScaleWidth      =   5625
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   5655
      Begin ComCtl2.Animation aniAnimation 
         Height          =   1095
         Left            =   2280
         TabIndex        =   7
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
         TabIndex        =   8
         Top             =   -240
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
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   855
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
      TabIndex        =   5
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
      Left            =   840
      ScaleHeight     =   8640
      ScaleWidth      =   10695
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   10695
      Begin MSFlexGridLib.MSFlexGrid flxVerordnung 
         Height          =   1455
         Left            =   9480
         TabIndex        =   41
         Top             =   5160
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.TextBox txtRezeptNr 
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   -120
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid flxKkassen 
         Height          =   1455
         Left            =   6600
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxRezeptDaten 
         Height          =   1455
         Left            =   4800
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
      End
      Begin VB.PictureBox picRezept 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5160
         ScaleHeight     =   1395
         ScaleWidth      =   3915
         TabIndex        =   19
         Top             =   3840
         Visible         =   0   'False
         Width           =   3975
         Begin VB.TextBox txtMarsRezeptAnmerkungen 
            Height          =   615
            Left            =   3120
            MultiLine       =   -1  'True
            TabIndex        =   37
            Text            =   "winrezk.frx":030A
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox cboVerfügbarkeit 
            Height          =   360
            Index           =   0
            Left            =   2400
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   30
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblMarsAnzahlRezepte 
            Alignment       =   2  'Zentriert
            BackColor       =   &H0000FF00&
            Caption         =   "Label1"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblWinVk 
            BackStyle       =   0  'Transparent
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
            Left            =   840
            TabIndex        =   34
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame fmeAv 
         Caption         =   "A+V Parameter"
         Height          =   2895
         Left            =   1200
         TabIndex        =   20
         Top             =   4440
         Width           =   2175
         Begin VB.CommandButton cmdAv 
            Cancel          =   -1  'True
            Caption         =   "A+&V"
            Height          =   450
            Left            =   360
            TabIndex        =   21
            Top             =   1800
            Width           =   1200
         End
         Begin VB.Label lblAv 
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
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblAv 
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
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblAv 
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
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdEsc 
         Caption         =   "ESC"
         Height          =   450
         Index           =   0
         Left            =   5280
         TabIndex        =   14
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid flxarbeit 
         Height          =   2760
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Tag             =   "0"
         Top             =   720
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   4868
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
         Interval        =   50
         Left            =   720
         Top             =   1080
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   540
         Index           =   0
         Left            =   600
         TabIndex        =   27
         Top             =   840
         Width           =   6720
         _ExtentX        =   11853
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
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfoZusatz 
         Height          =   780
         Index           =   0
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
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
      Begin VB.Label lblRezeptNr 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rezept-Nr."
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
         Left            =   0
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblMarsModus 
         Alignment       =   2  'Zentriert
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
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
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
         Left            =   120
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   270
         Width           =   9615
      End
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
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   4
      Left            =   12720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":0310
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1762
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":47B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":5C06
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":7058
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":84AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":98FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":AD4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":DDA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":F1F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":10644
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":11A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":12EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1433A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1578C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":16BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":18030
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1B082
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1C4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1D926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   3
      Left            =   12720
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1ED78
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":1FACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":22B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2386E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":245C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":25312
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":26064
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":26DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":29E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2AB5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2B8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2C5FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2D350
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2E0A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2EDF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":2FB46
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":30898
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":338EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3463C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3538E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   2
      Left            =   12720
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":360E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":36932
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":39984
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3A1D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3AA28
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3B27A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3BACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3C31E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3F370
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":3FBC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":40414
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":40C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":414B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":41D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4255C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":42DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":43600
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":46652
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":46EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":476F6
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
         NumListImages   =   25
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":47F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":481DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":484F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":48786
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":48A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":48CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":48FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":490D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":493F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":49682
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4999C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":49CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":49FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4A262
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4A57C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4A896
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4ABB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4AECA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4B15C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4B3EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4B680
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4B99A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4BCB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4BFCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4C2E8
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
            Picture         =   "winrezk.frx":4C602
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4C714
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4C9A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4CAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4CBCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4CE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4D0EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4D200
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4D51A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4D7AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4DAC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4DDE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4E0FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4E38C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4E6A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4E9C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4ECDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4EFF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4F106
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4F218
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4F32A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4F644
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4F95E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4FC78
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "winrezk.frx":4FF92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "&Datei"
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Rezept&speicher"
         Index           =   0
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "&Importkontrolle"
         Index           =   1
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "&Hilfsmittel-Nummern"
         Index           =   2
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Sonder-&Pzn"
         Index           =   3
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Kassen-Au&fschlag"
         Index           =   4
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Abholer-&Verwaltung"
         Index           =   5
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "&AVP Sendemodul"
         Index           =   6
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Five&Rx-Übertragung"
         Index           =   7
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "eRe&zepte"
         Index           =   8
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Selbst&erklärung"
         Index           =   9
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Schutz&masken"
         Index           =   10
         Begin VB.Menu mnuSchutzmaskenInd 
            Caption         =   "Schutzmaskenset Coupon 1"
            Index           =   0
         End
         Begin VB.Menu mnuSchutzmaskenInd 
            Caption         =   "Schutzmaskenset Coupon 2"
            Index           =   1
         End
         Begin VB.Menu mnuSchutzmaskenInd 
            Caption         =   "Schutzmaskenset Coupon ALG 2"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Pharm. D&L"
         Index           =   11
         Begin VB.Menu mnuPharmDienstleistungenInd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "Impfleis&tungen"
         Index           =   12
         Begin VB.Menu mnuImpfLeistungenInd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuDateiInd 
         Caption         =   "PflegeHilfsmittel-Abrechnung"
         Index           =   13
      End
      Begin VB.Menu mnuDummy17 
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
         Caption         =   ""
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "&Inhaltsstoffe"
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
         Caption         =   "&Automatische Rezeptauswahl"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Rezept-&Anmerkungen"
         Enabled         =   0   'False
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
         Caption         =   "Impfstoff-Rezept"
         Index           =   9
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Neu - Privatrezept"
         Index           =   10
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "A+V Parameter"
         Index           =   11
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "S&tatistik"
         Index           =   12
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Neu - zuz.frei"
         Index           =   13
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Neu - zuz.pflichtig"
         Index           =   14
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "Indiv. Rezeptur"
         Index           =   15
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuBearbeitenInd 
         Caption         =   "A+V Debug"
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
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Barverkäufe"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Ident"
         Index           =   1
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Import"
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "ABDA"
         Index           =   3
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Rezept - zuz.frei"
         Index           =   4
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Rezept - zuz.pflichtig"
         Index           =   5
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Stammdaten Kunden"
         Index           =   6
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "PlusX"
         Index           =   7
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuBearbeitenZusatz 
         Caption         =   "Nicht-Verfügbarkeit"
         Index           =   8
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDummy14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrivRezVK 
         Caption         =   "Privatrezepte holen"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuDummy15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLennartz 
         Caption         =   "Import Lennartz-Rezepturen"
         Shortcut        =   ^{F8}
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
      Begin VB.Menu mnuFarbe 
         Caption         =   "Farbe &Arbeitsbereich ..."
         Index           =   0
      End
      Begin VB.Menu mnuFarbe 
         Caption         =   "Farbe &Infobereich ..."
         Index           =   1
      End
      Begin VB.Menu mnuRezeptfarben 
         Caption         =   "Farbe &Rezepte"
         Begin VB.Menu mnuRezeptFarbenInd 
            Caption         =   "&Kassen-Rezept ..."
            Index           =   0
         End
         Begin VB.Menu mnuRezeptFarbenInd 
            Caption         =   "&Privat-Rezept ..."
            Index           =   1
         End
         Begin VB.Menu mnuRezeptFarbenInd 
            Caption         =   "&BTM-Rezept ..."
            Index           =   2
         End
         Begin VB.Menu mnuRezeptFarbenInd 
            Caption         =   "&Sonder-Belege ..."
            Index           =   3
         End
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
   End
   Begin VB.Menu mnuExtras 
      Caption         =   "E&xtras"
      Begin VB.Menu mnuOptionen 
         Caption         =   "&Optionen ..."
      End
      Begin VB.Menu mnuDummy4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrucker 
         Caption         =   "&Drucker"
         Begin VB.Menu mnuDruckerInd 
            Caption         =   "&Ohne Windows-Treiber"
            Index           =   0
            Begin VB.Menu MnuDosDruckerInd 
               Caption         =   "&Parameter ..."
               Index           =   0
            End
            Begin VB.Menu MnuDosDruckerInd 
               Caption         =   "-"
               Index           =   1
            End
         End
         Begin VB.Menu mnuDruckerInd 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuDruckerInd 
            Caption         =   "Windows-&Parameter"
            Index           =   2
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "Versatz&X ..."
               Index           =   0
            End
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "Versatz&Y ..."
               Index           =   1
            End
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "VersatzX &Datum ..."
               Index           =   2
            End
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "VersatzY &Datum ..."
               Index           =   3
            End
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "VersatzY &RezeptNr..."
               Index           =   4
            End
            Begin VB.Menu mnuDruckerWinPara 
               Caption         =   "&Schriftart ..."
               Index           =   5
            End
            Begin VB.Menu mnuDummy3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuCode128 
               Caption         =   "&Code 128"
               Begin VB.Menu mnuCode128Ind 
                  Caption         =   "Versatz&X ..."
                  Index           =   0
               End
               Begin VB.Menu mnuCode128Ind 
                  Caption         =   "Versatz&Y ..."
                  Index           =   1
               End
               Begin VB.Menu mnuCode128Ind 
                  Caption         =   "&Schriftart ..."
                  Index           =   2
               End
            End
         End
      End
      Begin VB.Menu mnuDummy6 
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

Const INI_SECTION = "Rezeptkontrolle"
Const INFO_SECTION = "Infobereich Rezeptkontrolle"


'Dim scrAuswahlAltValue%
Dim InRowColChange%

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

Dim HochfahrenAktiv%

Dim Standard%

Dim ClickOk%

Private Const DefErrModul = "WINREZK.FRM"

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
Static MnuEnabled%(16)
Static WarWechsel%

Select Case NeuerModus%
    Case 0
        mnuDatei.Enabled = True
        mnuBearbeiten.Enabled = True
        mnuAnsicht.Enabled = True
        mnuExtras.Enabled = True
        
        If (WarWechsel%) Then
            For i% = 0 To 16
                mnuBearbeitenInd(i%).Enabled = MnuEnabled%(i%)
            Next i%
        Else
            mnuBearbeitenInd(MENU_F2).Enabled = True
            mnuBearbeitenInd(MENU_F3).Enabled = True
            mnuBearbeitenInd(MENU_F4).Enabled = True
            mnuBearbeitenInd(MENU_F5).Enabled = True
            mnuBearbeitenInd(MENU_F6).Enabled = False
            mnuBearbeitenInd(MENU_F7).Enabled = False
            mnuBearbeitenInd(MENU_F8).Enabled = False
            mnuBearbeitenInd(MENU_F9).Enabled = True
            mnuBearbeitenInd(MENU_SF2).Enabled = ImpfstoffeDa%
            mnuBearbeitenInd(MENU_SF3).Enabled = True
            mnuBearbeitenInd(MENU_SF4).Enabled = VmFlag%
            mnuBearbeitenInd(MENU_SF5).Enabled = True
            mnuBearbeitenInd(MENU_SF6).Enabled = True
            mnuBearbeitenInd(MENU_SF7).Enabled = True
            mnuBearbeitenInd(MENU_SF8).Enabled = True
            mnuBearbeitenInd(MENU_SF9).Enabled = VmFlag%
        End If
        
        mnuBearbeitenLayout.Checked = False
        
        cmdOk(0).Default = True
        cmdEsc(0).Cancel = True

        If (para.Newline) Then
            flxarbeit(0).BackColorSel = RGB(135, 61, 52)
            flxInfo(0).BackColorSel = RGB(135, 61, 52)
        Else
            flxarbeit(0).BackColorSel = vbHighlight
            flxInfo(0).BackColorSel = vbHighlight
        End If
        
'        If (ProgrammChar$ = "B") Then tmrAction.Enabled = True
        
        h$ = Me.Caption
        ind% = InStr(h$, " (EDITIER-MODUS)")
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        Me.Caption = h$

        If (Chef = False) Then
            Call MarsNoChef
        End If
        If (MarsModus = MARS_REZEPT_KONTROLLE) Then
            mnuBearbeitenInd(MENU_F7).Enabled = True
            mnuBearbeitenInd(MENU_F8).Enabled = True
        End If
    Case 1
        For i% = 0 To 16
            MnuEnabled%(i%) = mnuBearbeitenInd(i%).Enabled
        Next i%
        WarWechsel% = True
        
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
        
'        tmrAction.Enabled = False
               
        h$ = Me.Caption
        Me.Caption = h$ + " (EDITIER-MODUS)"
        
        flxInfo(0).row = 0
        flxInfo(0).col = 1
        flxInfo(0).SetFocus
End Select

For i% = 0 To 7
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i%).Enabled
Next i%
For i% = 8 To 15
    cmdToolbar(i% + 1).Enabled = mnuBearbeitenInd(i% + 1).Enabled
Next i%
Call ShowToolbar

'Me.Caption = ProgrammName$ + lblArbeit(NeuerModus%).Caption
        
ProgrammModus% = NeuerModus%

Call DefErrPop
End Sub

Private Sub cboVerfügbarkeit_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboVerfügbarkeit_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iind As Integer

iind = cboVerfügbarkeit(index).ListIndex

If (iind >= 1) And (iind <= 3) Then
    If (Trim(RezNr) = "") Then
        Call ActProgram.MKdurchKK(index)
    End If
ElseIf (iind = 6) Then
    Call ActProgram.WunschArtikel(index)
End If
Call ActProgram.CalcRezeptWerte

Call DefErrPop
End Sub

Private Sub cboVerfügbarkeit_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboVerfügbarkeit_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'Debug.Print Str(index) + Str(cboVerfügbarkeit(index).ListIndex)
'If (cboVerfügbarkeit(index).ListIndex = 6) Then
'    Call ActProgram.WunschArtikel(index)
'End If

Call DefErrPop
End Sub

Private Sub cmdDatei_Click(index As Integer)
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
Dim erg%, iBenutzerNr%, iOk%
Dim TaskId&

If (mnuDateiInd(index).Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

If (index = 0) Then
'    iBenutzerNr% = HoleBenutzerSignatur
'    If (iBenutzerNr = 1) Then
        If (picRezept.Visible) Then cmdEsc(0).Value = True
        RezSpeicherModus% = 0
        frmRezSpeicher.Show 1
'    Else
'        Call MessageBox("Problem: Aufruf des Rezeptspeichers nur mit Chef-Passwort möglich!", vbCritical)
'        Call DefErrPop: Exit Sub
'    End If
ElseIf (index = 1) Then
    If (Taetigkeiten(1).pers(0) > 0) Then
        iBenutzerNr% = HoleBenutzerSignatur
        iOk = 0
        If (iBenutzerNr% > 0) Then
            ActBenutzer = iBenutzerNr
            Call PruefeTaetigkeiten
            iOk = DarfImportKontrolle
        End If
        If (iOk = 0) Then
            Call MessageBox("Problem: Keine Berechtigung für den Aufruf der Importkontrolle!", vbCritical)
            Call DefErrPop: Exit Sub
        End If
    End If
    
    RezSpeicherModus% = 1
    frmRezSpeicher.Show 1
ElseIf (index = 2) Then
    Call ActProgram.EditHmNummern
ElseIf (index = 3) Then
    Call ActProgram.EditSonderPzn
ElseIf (index = 4) Then
    Call ActProgram.EditF7Kassen
ElseIf (index = 5) Then
    TaskId& = Shell("winabhol.exe", vbNormalFocus)
ElseIf (index = 6) Then
    TaskId& = Shell("AvpSend.exe", vbNormalFocus)
ElseIf (index = 7) Then
    TaskId& = Shell("FiveRx.exe", vbNormalFocus)
ElseIf (index = 8) Then
''    TaskId& = Shell("TI_BACK.exe", vbNormalFocus)
    txtRezeptNr.text = ""
    frm_eRezepte.Show 1
    If (txtRezeptNr.text <> "") Then
        cmdOk(0).Value = True
    End If
ElseIf (index = 9) Then
    Call ActProgram.SelbsterklaerungDruck
'ElseIf (Index = 10) Then
'    Call ActProgram.SchutzMaskenDruck
ElseIf (index = 13) Then
    Call PflegeHilfsmittelAbrechnung
End If

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
Dim x&

With picRezept
    If (.Visible) Then
        .Visible = False
        picBack(0).FillColor = wpara.FarbeArbeit
        picBack(0).Line (.Left - Screen.TwipsPerPixelX * 10, .Top - Screen.TwipsPerPixelY * 10)-(picBack(0).Width - 5, .Top + .Height + Screen.TwipsPerPixelX * 10), wpara.FarbeArbeit, BF
    End If
End With
With flxarbeit(0)
    picBack(0).FillColor = RGB(232, 217, 172)
    x = txtRezeptNr.Left + txtRezeptNr.Width + 300
    RoundRect picBack(0).hdc, (.Left) / Screen.TwipsPerPixelX, (.Top / Screen.TwipsPerPixelY) - 5, (.Left + x) / Screen.TwipsPerPixelX, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
    Call wpara.FillGradient(picBack(0), (.Left) / Screen.TwipsPerPixelX + 1, ((.Top + .Height / 2) / Screen.TwipsPerPixelY), (.Left + x) / Screen.TwipsPerPixelX - 2, (.Top + .Height) / Screen.TwipsPerPixelY - 60, RGB(232, 217, 172), RGB(242, 227, 182))
End With
With picBack(0)
    .ForeColor = RGB(180, 180, 180) ' vbWhite
    .FillStyle = vbSolid
    .FillColor = vbWhite
    RoundRect .hdc, (txtRezeptNr.Left - 60) / Screen.TwipsPerPixelX, (txtRezeptNr.Top - 30) / Screen.TwipsPerPixelY, (txtRezeptNr.Left + txtRezeptNr.Width + 60) / Screen.TwipsPerPixelX, (txtRezeptNr.Top + txtRezeptNr.Height + 15) / Screen.TwipsPerPixelY, 10, 10

    .Refresh
End With


flxRezeptDaten.Visible = False
flxKkassen.Visible = False
txtRezeptNr.SetFocus
flxVerordnung.Visible = False
Call ResetToolbar

Call MarsSaveRezeptAnmerkungen
If (MarsAutomaticModus) Then
    Call MarsNextRezept(True)
End If

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

Private Sub flxarbeit_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
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
    Call ActProgram.flxArbeitKeyPress(KeyAscii)
End If

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
Dim i%, ArbeitRow%, InfoRow%, aRow%, aCol%
Dim h$

'If (Index = 0) Then
'    ArbeitRow% = flxarbeit(0).row
'
'    With flxInfo(0)
'        aRow% = .row
'        aCol% = .col
'        InfoRow% = 0
'
'        Call ActProgram.flxInfoGotFocus(InfoRow%)
'
'        For i% = InfoRow% To (.Rows - 1)
'            .TextMatrix(i%, 0) = ""
'        Next i%
'        For i% = 0 To (.Rows - 1)
'            .row = i%
'            .col = 0
'            .CellFontBold = True
'            .CellForeColor = .ForeColor
'        Next i%
'
'        .row = aRow%
'        .col = aCol%
'    End With
'End If

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

If (KeinRowColChange% = False) Then
'    Error 6
    Call EchtKurzInfo
End If
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
    
Call DefErrPop
End Sub

Private Sub flxKkassen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKkassen_MouseMove")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

h = ""
On Error Resume Next
With flxKkassen
    ind = y \ .RowHeight(0) + .TopRow
    h = .TextMatrix(ind, 1)
End With
On Error GoTo DefErr

'flxInfo(0).TextMatrix(2, 0) = h

If (h <> "") Then
    With picToolTip
        .Width = .TextWidth(h$ + "x")
        .Height = .TextHeight(h$) + 45
        .Left = x + flxKkassen.Left + 300
        .Top = y + picBack(0).Top + flxKkassen.Top + 150
        .Visible = True
        .Cls
        .CurrentX = 2 * Screen.TwipsPerPixelX
        .CurrentY = 0
        picToolTip.Print h$
    End With
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
   

Call wpara.InitEndSub(Me)
Set opToolbar = New clsToolbar

Call wpara.HoleGlobalIniWerte(UserSection$, INI_DATEI, "WinRezK")
Call wpara.InitFont(Me)
Call HoleIniWerte

'ProgrammTyp% = Standard%
mnuDateiInd(6).Enabled = AvpTeilnahme%

Set InfoMain = New clsInfoBereich
Set opBereich = New clsOpBereiche

For i = 1 To 7
    Load mnuPharmDienstleistungenInd(i)
Next i
For i = 0 To 7
    mnuPharmDienstleistungenInd(i).Caption = PharmDienstleistungenTxt(i)
Next i

For i = 1 To 1
    Load mnuImpfLeistungenInd(i)
Next i
For i = 0 To 1
    mnuImpfLeistungenInd(i).Caption = ImpfLeistungenTxt(i)
Next i


Call InitDateiButtons

Call InitAnimation

Call InitProgrammTyp

If (para.Newline) Then
    mnuToolbar.Visible = False
    mnuFarbe(0).Visible = False
    mnuFarbe(1).Visible = False
    mnuFont(0).Caption = "Schriftart und -größe ..."
    mnuFont(1).Visible = False
    
    mnuNlToolbar.Visible = True
    mnuNlToolbarInd(opToolbar.Size).Checked = True

    mnuRezeptfarben.Visible = False
    mnuDummy7.Visible = False
End If

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
For i% = 1 To 8
    Unload mnuBearbeitenZusatz(i%)
Next i%
On Error GoTo DefErr

If (ProgrammTyp% = 0) Then
    Set ActProgram = New clsWinRezK
ElseIf (ProgrammTyp% = 2) Then
'    Set ActProgram = New clsWarenÜber
End If

Call ActProgram.Init(Me, opToolbar, InfoMain, opBereich)
ErstAuslesen% = True

picBack(0).Visible = True

Call DefErrPop
End Sub

Sub MarsRezeptKontrolleIcon()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MarsRezeptKontrolleIcon")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call opToolbar.InitToolbar(Me, App.EXEName, INI_SECTION, ",0518")

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
Call ProgrammEnde
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
Dim erg%, row%, col%, ind%
Dim l&
Dim h$, mErg$

If (para.Newline) Then
    ind = index
    If (ind <= MENU_F9) Then
        ind = ind + 1
    End If
    Call opToolbar.Click(-ind)
End If

Select Case index

    Case MENU_F2
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call EchtKurzInfo
            End If
        Else
            Call ActProgram.MenuBearbeiten(index)
        End If
    
    Case MENU_F3
'        If (ProgrammChar$ = "B") Then
'            BestellAnzeige% = (BestellAnzeige% + 1) Mod 3
'            Standard% = BestellAnzeige%
'            l& = WritePrivateProfileString(INI_SECTION, "Standard", Str$(Standard%), INI_DATEI)
'            BekartCounter% = -1
'            Call ActProgram.AuslesenBestellung(True, False, True)
'        End If
    
    Case MENU_F4
        Call ActProgram.MenuBearbeiten(index)
        
    Case MENU_F5
        If (ProgrammModus% = 1) Then
            If (ActiveControl.Name = flxInfo(0).Name) Then
                Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
                opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
                Call opBereich.RefreshBereich
                Call EchtKurzInfo
            End If
        Else
            Call ActProgram.MenuBearbeiten(index)
        End If
        
    Case MENU_F6
        If (picRezept.Visible) And (tmrF6Sperre.Enabled = False) Then
            Call ActProgram.MenuBearbeiten(index)
            tmrF6Sperre.Enabled = True
        End If
    
    Case MENU_F7
        Call frmAction.MarsInitAutomatikModus
        
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
'            Call ActProgram.MenuBearbeiten(index)
            With txtMarsRezeptAnmerkungen
                .Visible = True
                .SetFocus
            End With
        End If
    
    Case MENU_SF2
        Call ActProgram.MenuBearbeiten(index)
        
    Case MENU_SF3
        Call ActProgram.MenuBearbeiten(index)
        
    Case MENU_SF4
        Call ActProgram.MenuBearbeiten(index)
            
    Case MENU_SF5
        Call ActProgram.MenuBearbeiten(index)
'        Call ZeigeStatistik
        
    Case MENU_SF6
        Call ActProgram.MenuBearbeiten(index)
        
    Case MENU_SF7
        Call ActProgram.MenuBearbeiten(index)

    Case MENU_SF8
        Call ActProgram.MenuBearbeiten(index)

    Case MENU_SF9
        Call ActProgram.MenuBearbeiten(index)
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

Private Sub mnuBearbeitenZusatz_Click(index As Integer)
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
Call ActProgram.mnuBearbeitenZusatzClick(index)
Call DefErrPop
End Sub

Private Sub mnuDateiInd_Click(index As Integer)
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
cmdDatei(index).Value = True
Call DefErrPop
End Sub

Private Sub mnuFarbe_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuFarbe_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

erg% = wpara.EditFarbe(dlg, index)
If (erg%) Then
    Call opBereich.ResizeWindow
End If

Call DefErrPop
End Sub

Private Sub mnuLennartz_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuLennartz_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        
Call ActProgram.Lennartz

Call DefErrPop
End Sub

Private Sub mnuPharmDienstleistungenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuPharmDienstleistungenInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Sonderbeleg As New Sonderbeleg_OP
Dim sVerordnung As String
Dim sTaskId As String

'Dim sErg As String
'Dim pzn As String
'Dim txt As String
'Dim KuNr As Long
'sErg = MatchCode(5, pzn, txt, True, False)
'If (sErg <> "") Then
'    KuNr = xVal(sErg)
'End If

'Dim erg As Integer
'sTaskId = "910.024.941.037.669.73"
'erg = HoleErgebnisse(sTaskId)
'MsgBox (CStr(erg))
'Call DefErrPop: Exit Sub

Dim h As String
Dim KundenNr As Long
h = MyInputBox("Kunden-Nummer: ", "Abrechnung Pharm.Dienstleistungen", "")
h$ = Trim(h)
If (Val(h) > 0) Then
    KundenNr = Val(h)
Else
    Call DefErrPop: Exit Sub
End If
If (ActProgram.HoleSonderbelegKundenInfo(KundenNr, Sonderbeleg) = False) Then
    Call MessageBox("Problem: Kunden-Nummer " + CStr(KundenNr) + " NICHT VERGEBEN !", vbInformation)
    Call DefErrPop: Exit Sub
End If

If (FDok = 0) Then
    Set FD_OP = New TI_Back.Fachdienst_OP
    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), True)) Then
        FDok = True
    End If
End If

With Sonderbeleg
    .pzn = PharmDienstleistungenPzn(index)  '"17716872"
    .Verordnungstext = PharmDienstleistungenTxt(index)  ' "Standardisierte Risikoerfassung bei Bluthochdruck-Patienten"
'
'    With .Kostentraeger
'        .Typ = "GKV"
'        .IK = "101575519"
'        .Name = "Techniker Krankenkasse"
'    End With
'
'    With .Patient
'        .Id = "T555558879"
'        .GeburtsDatum = CDate("01.12.1999")
'        With .Name
'            .Vorname = "Andreas"
'            .Nachname_ohne_Vor_und_Zusatz = "Eder"
'            .Titel = "Dipl-Ing"
'        End With
'        With .Adresse
'            .PLZ = "1210"
'            .Ort = "Wien"
'            With .Strasse1
'                .Strasse = "Leopoldauer Strasse"
'                .Hausnummer = "68A 3 25"
'            End With
'        End With
'    End With
End With

sTaskId = FD_OP.TI_SonderBeleg(Sonderbeleg)

If (sTaskId <> "") Then
    Dim bErgebnis As Boolean

    Dim eRezept As New TI_Back.eRezept_OP
    Call eRezept.New2(sTaskId)
                            
    If (FD_OP.TI_RezeptAbrechnen(eRezept, False)) Then
'                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
        bErgebnis = VorabPruefung(sTaskId, False)
        If (bErgebnis) Then
        End If
    Else
        If (bEinzelErgebnis) Then
            Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
        End If
    End If
End If
        
Call DefErrPop
End Sub

Private Sub mnuImpfLeistungenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuImpfLeistungenInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Sonderbeleg As New Sonderbeleg_OP
Dim sVerordnung As String
Dim sTaskId As String

Dim h As String
Dim KundenNr As Long
h = MyInputBox("Kunden-Nummer: ", "Abrechnung Impfleistungen", "")
h$ = Trim(h)
If (Val(h) > 0) Then
    KundenNr = Val(h)
Else
    Call DefErrPop: Exit Sub
End If
If (ActProgram.HoleSonderbelegKundenInfo(KundenNr, Sonderbeleg) = False) Then
    Call MessageBox("Problem: Kunden-Nummer " + CStr(KundenNr) + " NICHT VERGEBEN !", vbInformation)
    Call DefErrPop: Exit Sub
End If

If (FDok = 0) Then
    Set FD_OP = New TI_Back.Fachdienst_OP
    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), True)) Then
        FDok = True
    End If
End If

ImpfLeistungModus = IIf(index = 0, "G", "C")
frmImpfPzns.Show 1
If (FormErg = 0) Then
    Call DefErrPop: Exit Sub
End If
'MsgBox (FormErgTxt)
Dim ImpfZeile() As String
ImpfZeile = Split(FormErgTxt, "-")

With Sonderbeleg
    Dim dGesamtBrutto As Double
    Dim iAnzDosen As Integer
    
    iAnzDosen = Val(ImpfZeile(3))
    
    .pzn = ImpfLeistungenPzn(index)  '"17716872"
    .Verordnungstext = ImpfLeistungenTxt(index)  ' "Standardisierte Risikoerfassung bei Bluthochdruck-Patienten"
    
    .FlowType = 920
    If (ImpfLeistungModus = "G") Then
        h = "S17716926-10,4;S17716955-0,7;S18774512-1;"
        dGesamtBrutto = 12.1
    Else
        h = "S17717400-10;"
        dGesamtBrutto = dGesamtBrutto + 10
        If (iAnzDosen > 1) Then
            If (MessageBox("Vergütung für den Umgang mit Mehrdosisbehältnissen bei Schutzimpfungen gegen das Coronavirus SARS-CoV-2 (PZN 17717417, 2.5 EUR) einfügen ?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Impfleistungen") = vbYes) Then
                h = h + "S17717417-2,5;"
                dGesamtBrutto = dGesamtBrutto + 2.5
            End If
        End If
        If (MessageBox("Vergütung für gegebenenfalls erforderlichen weiteren Aufwand (PZN 17717423, 2.5 EUR) einfügen ?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Impfleistungen") = vbYes) Then
            h = h + "S17717423-2,5;"
            dGesamtBrutto = dGesamtBrutto + 2.5
        End If
    End If
    
    h = h + IIf(iAnzDosen > 1, "S02567053", ImpfZeile(0)) + "-" + ImpfZeile(2) + "-" + ImpfZeile(3) + ";"
    dGesamtBrutto = dGesamtBrutto + xVal(ImpfZeile(2))
    
    .ImpfleistungPZNs = h
    .sGesamtBrutto = uFormat(dGesamtBrutto, "0.00")
    
    If (iAnzDosen > 1) Then
        ParEnteralHerstellerKey = 3
        With .iHerstellung
            .Charge = 0 ' "zz"
            .HerstellerKeyTA3 = CStr(ParEnteralHerstellerKey)
            .HerstellerKz = Right(String(9, "0") + ParEnteralHerstellerKz(ParEnteralHerstellerKey), 9)
            .Herstellungszeit = Format(Now, "YYYYMMDD:0000")
        End With
            
        TA1_V37 = 1
        With .iBestandteil
            .Artikelnummer = ImpfZeile(0)
            .ArtikelnummerTyp = 1   'bei ADV: PZN   'AvpRec!flag
            .Faktor = HashFaktor(1 / iAnzDosen * 1000)
            .Faktorkennzeichen = "11"
            .Preiskennzeichen = "41"
            .Taxe = FiveRxPreis(xVal(ImpfZeile(1)))
            .Positionsnummer = 1
        End With
    End If
    
    .Typ = IIf(ImpfLeistungModus = "G", "Grippe-Impfleistung", "Corona-Impfleistung")
End With

sTaskId = FD_OP.TI_SonderBeleg(Sonderbeleg)

If (sTaskId <> "") Then
    Dim bErgebnis As Boolean

    Dim eRezept As New TI_Back.eRezept_OP
    Call eRezept.New2(sTaskId)
'    eRezept.ImpfleistungPZNs = Sonderbeleg.ImpfleistungPZNs
                            
    If (FD_OP.TI_RezeptAbrechnen(eRezept, False)) Then
'                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
        bErgebnis = VorabPruefung(sTaskId, False)
        If (bErgebnis) Then
        End If
    Else
        If (bEinzelErgebnis) Then
            Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
        End If
    End If
End If
        
Call DefErrPop
End Sub

Private Sub PflegeHilfsmittelAbrechnung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PflegeHilfsmittelAbrechnung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Sonderbeleg As New Sonderbeleg_OP
Dim sVerordnung As String
Dim sTaskId As String
'Dim PflegeHilfsmittelNummern As String

Dim h As String
Dim KundenNr As Long

'PflegeHilfsmittelNummern = HolePflegeHilfsmittelNummern

h = MyInputBox("Kunden-Nummer: ", "Abrechnung Pflegehilfsmittel", "")
h$ = Trim(h)
If (Val(h) > 0) Then
    KundenNr = Val(h)
Else
    Call DefErrPop: Exit Sub
End If
If (ActProgram.HoleSonderbelegKundenInfo(KundenNr, Sonderbeleg) = False) Then
    Call MessageBox("Problem: Kunden-Nummer " + CStr(KundenNr) + " NICHT VERGEBEN !", vbInformation)
    Call DefErrPop: Exit Sub
End If
'
If (FDok = 0) Then
    Set FD_OP = New TI_Back.Fachdienst_OP
    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), (wpara.EntwicklungsUmgebung = 0))) Then
        FDok = True
    End If
End If

Dim Ti1 As New TI_PflegeHilfmittel.clsPflegeHilfmittel_OP
Set Sonderbeleg = Ti1.SonderbelegPflegeHilfsmittel(KundenNr, "")
''MsgBox (Ti1.PflegeHilfsmittel(Val(h), Me))
Set Ti1 = Nothing
'
If (xVal(Sonderbeleg.sGesamtBrutto) <= 0) Then
    Call MsgBox("Problem: KEINE PflegeHilfsmittel ausgewählt !", vbInformation)
    Call DefErrPop: Exit Sub
End If

Sonderbeleg.VerordnungsDatum = CDate("01." + Format(Now, "MM.yyyy"))
    
'If (fdok = 0) Then
'    Set FD_OP = New TI_Back.Fachdienst_OP
'    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), (wpara.EntwicklungsUmgebung = 0))) Then
'        fdok = True
'    End If
'End If
'
sTaskId = FD_OP.TI_SonderBeleg(Sonderbeleg)

If (sTaskId <> "") Then
End If
Set Sonderbeleg = Nothing
        
Call DefErrPop
End Sub

'Private Sub mnuPharmDienstleistungenInd_Click(index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("mnuPharmDienstleistungenInd_Click")
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
'Dim Sonderbeleg As New Sonderbeleg_OP
'Dim sVerordnung As String
'Dim sTaskId As String
'
''Dim sErg As String
''Dim pzn As String
''Dim txt As String
''Dim KuNr As Long
''sErg = MatchCode(5, pzn, txt, True, False)
''If (sErg <> "") Then
''    KuNr = xVal(sErg)
''End If
'
''Dim erg As Integer
''sTaskId = "910.024.941.037.669.73"
''erg = HoleErgebnisse(sTaskId)
''MsgBox (CStr(erg))
''Call DefErrPop: Exit Sub
'
'Dim h As String
'Dim KundenNr As Long
'h = MyInputBox("Kunden-Nummer: ", "Abrechnung Pharm.Dienstleistungen", "")
'h$ = Trim(h)
'If (Val(h) > 0) Then
'    KundenNr = Val(h)
'Else
'    Call DefErrPop: Exit Sub
'End If
'If (ActProgram.HoleSonderbelegKundenInfo(KundenNr, Sonderbeleg) = False) Then
'    Call MessageBox("Problem: Kunden-Nummer " + CStr(KundenNr) + " NICHT VERGEBEN !", vbInformation)
'    Call DefErrPop: Exit Sub
'End If
'
'If (FDok = 0) Then
'    Set FD_OP = New TI_Back.Fachdienst_OP
'    If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd), True)) Then
'        FDok = True
'    End If
'End If
'
'With Sonderbeleg
'    .pzn = PharmDienstleistungenPzn(index)  '"17716872"
'    .Verordnungstext = PharmDienstleistungenTxt(index)  ' "Standardisierte Risikoerfassung bei Bluthochdruck-Patienten"
''
''    With .Kostentraeger
''        .Typ = "GKV"
''        .IK = "101575519"
''        .Name = "Techniker Krankenkasse"
''    End With
''
''    With .Patient
''        .Id = "T555558879"
''        .GeburtsDatum = CDate("01.12.1999")
''        With .Name
''            .Vorname = "Andreas"
''            .Nachname_ohne_Vor_und_Zusatz = "Eder"
''            .Titel = "Dipl-Ing"
''        End With
''        With .Adresse
''            .PLZ = "1210"
''            .Ort = "Wien"
''            With .Strasse1
''                .Strasse = "Leopoldauer Strasse"
''                .Hausnummer = "68A 3 25"
''            End With
''        End With
''    End With
'End With
'
'sTaskId = FD_OP.TI_SonderBeleg(Sonderbeleg)
'
'If (sTaskId <> "") Then
'    Dim bErgebnis As Boolean
'
'    Dim eRezept As New TI_Back.eRezept_OP
'    Call eRezept.New2(sTaskId)
'
'    If (FD_OP.TI_RezeptAbrechnen(eRezept, False)) Then
''                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
'        bErgebnis = VorabPruefung(sTaskId, False)
'        If (bErgebnis) Then
'        End If
'    Else
'        If (bEinzelErgebnis) Then
'            Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
'        End If
'    End If
'End If
'
'Call DefErrPop
'End Sub

Private Sub mnuPrivRezVK_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuPrivRezVK_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
frmFortschritt.Show vbModal
Call DefErrPop
End Sub

Private Sub mnuRezeptFarbenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuRezeptFarbenInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim l&
Dim Key$

erg% = EditFarbe(index)
If (erg%) Then
    Key$ = "FarbeRezept" + Format(index, "0")
    l& = WritePrivateProfileString("Rezeptkontrolle", Key$, Hex$(RezeptFarben&(index)), INI_DATEI)
    If (picRezept.Visible) Then ActProgram.PaintRezept
End If

Call DefErrPop
End Sub

Function EditFarbe%(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditFarbe%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%, ret%
Dim l&, lColor&

On Error Resume Next

ret% = False

With frmAction.dlg
    .color = RezeptFarben&(index)
    .CancelError = True
    .flags = cdlCCFullOpen + cdlCCRGBInit
    Call .ShowColor
    If (Err = 0) Then
        RezeptFarben&(index) = .color
        ret% = True
    End If
End With

EditFarbe% = ret%

Call DefErrPop
End Function

Function EditFont%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditFont%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

ret% = False

With frmAction.dlg
    .FontName = RTrim$(RezeptFont$)
    .FontSize = 12
    .FontBold = False
    .FontItalic = False
    .FontStrikethru = False
    .FontUnderline = False
    .flags = cdlCFBoth
    .CancelError = False
    Call frmAction.dlg.ShowFont
    If (Err = 0) Then
        If (RezeptFont$ <> .FontName) Then
            RezeptFont$ = .FontName
            ret% = True
        End If
    End If
End With

EditFont% = ret%

Call DefErrPop
End Function

Function EditCode128Font%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditCode128Font%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

ret% = False

With frmAction.dlg
    .FontName = RTrim$(Code128Font$)
    .FontSize = Code128FontSize
    .FontBold = False
    .FontItalic = False
    .FontStrikethru = False
    .FontUnderline = False
    .flags = cdlCFBoth
    .CancelError = False
    Call frmAction.dlg.ShowFont
    If (Err = 0) Then
'        If (RezeptFont$ <> .FontName) Then
            Code128Font$ = .FontName
            Code128FontSize = .FontSize
            ret% = True
'        End If
    End If
End With

EditCode128Font% = ret%

Call DefErrPop
End Function

Private Sub mnuFont_Click(index As Integer)
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

erg% = wpara.EditFont(dlg, index)
If (erg%) Then
    Call wpara.InitFont(frmAction)
    Call opBereich.ResizeWindow

    frmAction.flxarbeit(0).Rows = 1
    frmAction.flxInfo(0).Clear
    Call ActProgram.mnuFontClick
End If

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

If (Shift And vbAltMask And (KeyCode = vbKeyG)) Then
    Call ActProgram.ImporteOrAutIdem(False)
    KeyCode = 0
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
'        Case vbKeyG
'            ind% = -2
        Case vbKeyS
            ind% = -1
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
    ElseIf (ind% = -1) Then
        h$ = "02000"
'        If ((BestellAnzeige% = 2) And (IstDirektLief%)) Then h$ = "02091"
'        Call Stammdaten(Format(Lieferant%, "000"), Val(h$), DirektBewertung#, ActProgram)
        Call ActProgram.MenuBearbeiten(MENU_F7)
        KeyCode = 0
'    ElseIf (ind% = -2) Then
'        Call ActProgram.ImporteOrAutIdem(False)
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

If (picRezept.Visible) Then Call ActProgram.PaintRezept

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

Private Sub mnuNlToolbarInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuNlToolbarInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

opToolbar.Size = index
For i = 0 To 2
    mnuNlToolbarInd(i).Checked = (i = index)
Next i

Call DefErrPop
End Sub

Private Sub mnuSchutzmaskenInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuSchutzmaskenInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call ActProgram.SchutzMaskenDruck(index + 1)

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

Private Sub picRezept_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picRezept_KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call ActProgram.KeyDown(KeyCode, Shift)
Call DefErrPop
End Sub

Private Sub picRezept_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picRezept_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'If (KeyAscii >= 48) And (KeyAscii <= 57) Then
'    EingabeStr$ = EingabeStr$ + Chr$(KeyAscii)
'End If

    Dim l%
    Dim ch$
    
    ch = Chr(KeyAscii)
    l = Len(EingabeStr)
    If (InStr("0123456789", ch) > 0) Then
        EingabeStr$ = EingabeStr$ + ch
    ElseIf (l = 0) Then
        If (ch = "[") Then
            EingabeStr$ = EingabeStr$ + ch
        End If
    ElseIf (l > 15) Then
        EingabeStr$ = EingabeStr$ + ch
    Else
        If (l > 3) Then
            l = 3
        End If
        If (Left(EingabeStr, l) = Left("[)>", l)) Then
            EingabeStr$ = EingabeStr$ + ch
        End If
    End If

Call DefErrPop
End Sub

Private Sub picRezept_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picRezept_MouseDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

ClickOk% = ActProgram.MouseDown(x, y)

Call DefErrPop
End Sub

Private Sub picRezept_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picRezept_DblClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (ClickOk%) Then Call ActProgram.EditSatz

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
            Call EchtKurzInfo
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
ElseIf (ActiveControl.Name = flxKkassen.Name) Then
    Call ActProgram.flxKkassenClick
ElseIf (ActiveControl.Name = picRezept.Name) Then
    If (Len(EingabeStr) > 15) Then
        EingabeStr = CheckSecurPharm(EingabeStr)
    End If
    
    Call ActProgram.EditSatz
    EingabeStr$ = ""
ElseIf (ActiveControl.Name = txtMarsRezeptAnmerkungen.Name) Then
Else
    RezNr$ = Trim(txtRezeptNr.text)
    If (ActProgram.RezeptHolen) Then
        Call ActProgram.PaintRezept(1)
        
        If (PreisDiffStr$ <> "") Then
            tmrAction.Enabled = True
            PreisDiffStr = "Folgende Artikel haben einen von der Taxe abweichenden Preis:" + vbCrLf + PreisDiffStr
            Call MsgBox2(Me.hWnd, PreisDiffStr, "Hinweis", vbOKOnly)
        End If
        
        Call ActProgram.RepaintBtmGebühr
        Call ActProgram.ShowNichtInTaxe
    Else
        Call txtRezeptNr_GotFocus
        txtRezeptNr.SetFocus
    End If
End If

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

If (eRezeptTaskId <> "") Then
    If (wbVerordnung.Visible) Then
        Dim sDatei$, sBundleDAV$
        Dim l&
        
        sDatei = CurDir() + "\eVerordnung.xml"
        l = writeOut(eRezeptBundleKBV, sDatei)
        
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
        
        If (FDok = 0) Then
            Set FD_OP = New TI_Back.Fachdienst_OP
            If (FD_OP.Init(FD_OP.NetWindowHandle(Me.hWnd))) Then
                FDok = True
            End If
        End If
        
        Dim sVersionAbgabedaten As String
        Dim eRezept As New TI_Back.eRezept_OP
        Call eRezept.New2(eRezeptTaskId)
        sBundleDAV = FD_OP.TI_RezeptAbgabedaten(eRezept, sVersionAbgabedaten)
'        MsgBox (sVersionAbgabedaten)
      
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
    End If

    Call DefErrPop: Exit Sub
    
    Dim bOk%, iAktion%, ind%
    Dim h$, sAktion$, sTag$
    flxInfo(0).Rows = flxInfo(0).FixedRows
    SQLStr = "SELECT * FROM TI_Aktionen"
    SQLStr = SQLStr + " WHERE (TaskId='" + eRezeptTaskId + "')"
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
        h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "YYMMDDHHmm") + Format(VerkaufAdoRec!Id, "000000")
        flxInfo(0).AddItem h, 1
        
        VerkaufAdoRec.MoveNext
    Loop
    VerkaufAdoRec.Close
    
    SQLStr = "SELECT * FROM TI_FiveRx" '
    SQLStr = SQLStr + " WHERE (TaskId='" + eRezeptTaskId + "')"
    SQLStr = SQLStr + " ORDER BY AnlageDatum"
    FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
    Do
        If (VerkaufAdoRec.EOF) Then
            Exit Do
        End If
        
        iAktion = CheckNullInt(VerkaufAdoRec!AktionInt)
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
            h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!AnlageDatum), "YYMMDDHHmm") + Format(VerkaufAdoRec!Id, "000000")
            flxInfo(0).AddItem h, 1
        Else
            sAktion = CheckNullStr(VerkaufAdoRec!FiveRxXml)
            flxInfo(0).TextMatrix(1, 5) = sAktion
            If (iAktion = 1) Or (iAktion = 5) Then
                sTag = "STATUS"
            Else
                sTag = "RZLIEFERID"
            End If
            ind = InStr(UCase(sAktion), "<" + sTag + ">")
            If (ind > 0) Then
                sAktion = Mid(sAktion, ind + Len(sTag) + 1 + 1)
                ind = InStr(UCase(sAktion), "</" + sTag + ">")
                If (ind > 0) Then
                    sAktion = Left(sAktion, ind - 1)
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
            .Sort = 4   'Zahlen absteigend
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
Else
    Call ActProgram.EchtKurzInfo
End If

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
Dim KalkAvp#, RundAvp#
Dim BekLaufNr&
Dim h$, KalkText$
Static aBekLaufNr&, aFlexRow%

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
        .HighLight = flexHighlightWithFocus
        KeinRowColChange% = False
    Else
        BekLaufNr& = Val(.TextMatrix(.row, 20))
        If (BekLaufNr& <> aBekLaufNr&) Or (aFlexRow% <> .row) Then
            
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
            
            If (wpara.FarbeAktZeile) Then
                .CellForeColor = vbHighlight
            Else
                .CellForeColor = vbHighlightText
            End If
            
            .FillStyle = flexFillSingle
            .col = aCol%
            
            Call EchtKurzInfo
            aBekLaufNr& = BekLaufNr&
            aFlexRow% = .row
            .HighLight = flexHighlightWithFocus
            KeinRowColChange% = False
            
            ZeilenWechsel% = True
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
Dim i%, br%
Dim x&

If (para.MARS) Then
    With lblMarsModus
        .Left = wpara.LinksX
        .Top = wpara.TitelY   'FlexY%
        .Width = txtRezeptNr.Width
        
'        ToolbarBackR = 135
'        ToolbarBackG = 61
'        ToolbarBackB = 52
'        ToolbarBackR = 201
'        ToolbarBackG = 123
'        ToolbarBackB = 58
    
        .FontSize = frmAction.Font.Size + 4
        .Height = TextHeight("Äg" + vbCrLf + "Äg" + vbCrLf) + 90
        
        .BackColor = RGB(201 + 20, 123 + 20, 58 + 20)
        .ForeColor = vbWhite    'RGB(210, 210, 210)
        
        .Visible = True
    End With
    With lblRezeptNr
        .Left = wpara.LinksX
        .Top = lblMarsModus.Top + lblMarsModus.Height + 600
    End With
Else
    With lblRezeptNr
        .Left = wpara.LinksX
        .Top = wpara.TitelY   'FlexY%
    End With
End If
With txtRezeptNr
    .Left = wpara.LinksX ' lblRezeptNr.Left + lblRezeptNr.Width + 150
    .Top = lblRezeptNr.Top + lblRezeptNr.Height + 45
    .Width = TextWidth(String(13, "9")) + 90
End With

For i% = 0 To 2
    With lblAv(i%)
        .Left = 90
        If (i% = 0) Then
            .Width = TextWidth(String(20, "X"))
            .Top = 2 * wpara.TitelY
        Else
            .Width = lblAv(i% - 1).Width
            .Top = lblAv(i% - 1).Top + .Height + 15
        End If
    End With
Next i%
With cmdAv
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (lblAv(0).Width - .Width) / 2
    .Top = lblAv(2).Top + lblAv(2).Height
End With
With fmeAv
    .Width = lblAv(0).Width + 2 * lblAv(0).Left
    .Height = cmdAv.Top + cmdAv.Height + 90
    .Left = wpara.LinksX
    .Top = txtRezeptNr.Top + txtRezeptNr.Height + 150   ' 300
    x& = .Left + .Width + 300
    .Visible = False
End With

'With flxRezeptDaten
'    .Width = fmeAv.Width
'    .Height = .Rows * .RowHeight(0) + 90
'    .Left = fmeAv.Left
'    .Top = fmeAv.Top + fmeAv.Height + 300
'    .TextMatrix(0, 0) = "Datum"
'    .TextMatrix(1, 0) = "Zeit"
'    .TextMatrix(2, 0) = "Ort"
'    .TextMatrix(3, 0) = "Benutzer"
'    .TextMatrix(5, 0) = "Gedruckt"
'    .ColWidth(0) = TextWidth("Benutzer") + 150
'    .ColWidth(1) = .Width - .ColWidth(0)
'    .ColAlignment(1) = flexAlignLeftCenter
'    .BackColor = wpara.FarbeArbeit
'    .GridColor = wpara.FarbeArbeit
'End With

With flxRezeptDaten
    .Visible = False
    .Rows = 2
    .Cols = 4
    .FixedCols = 0
    .Height = .Rows * .RowHeight(0) + 90
    
'    .Width = fmeAv.Width
'    .Height = .Rows * .RowHeight(0) + 90
'    .Left = fmeAv.Left
'    .Top = fmeAv.Top + fmeAv.Height + 300
    .TextMatrix(0, 0) = "Abgabe"
    .TextMatrix(1, 0) = "Benutzer"
    .TextMatrix(0, 2) = "Gedruckt"
    
    .ColWidth(0) = TextWidth("Benutzer") + 150
    .ColWidth(1) = TextWidth("999999  99:99   Comp:99R") + 150
    .ColWidth(2) = TextWidth("Gedruckt") + 150
    .ColWidth(3) = TextWidth(String(8, "9"))
    
    For i% = 0 To (.Cols - 1)
        br% = br% + .ColWidth(i%)
    Next i%
    .Width = br% + 90
    
    For i% = 0 To (.Cols - 1) Step 2
        .ColAlignment(i%) = flexAlignLeftCenter
    Next i%
    For i% = 1 To (.Cols - 1) Step 2
        .ColAlignment(i%) = flexAlignRightCenter
        .FillStyle = flexFillRepeat
        .row = 0
        .col = i%
        .RowSel = .Rows - 1
        .ColSel = .col
        .CellForeColor = vbWhite
        .FillStyle = flexFillSingle
    Next i%
    .BackColor = wpara.FarbeArbeit
    .GridColor = wpara.FarbeArbeit
End With

With flxKkassen
    If (.Rows = 0) Then .Rows = 1
    .Left = fmeAv.Left
    .Top = fmeAv.Top '+ fmeAv.Height + 150
    .Width = fmeAv.Width
    
    .Height = opBereich.ArbeitBackHeight - .Top - 150
    .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
    
    .ColAlignment(0) = flexAlignLeftCenter
    .ColWidth(0) = 0
    .ColWidth(1) = .Width
    .ColWidth(2) = 0
    .ColWidth(3) = 0
    .BackColor = wpara.FarbeArbeit
    .BackColorBkg = wpara.FarbeArbeit
    .GridColor = wpara.FarbeArbeit

    lblMarsModus.Width = .Width
End With
        
With picRezept
    .Font.Name = wpara.FontName(0)
    .Font.Size = wpara.FontSize(0)
    If ((ScaleWidth - .Width) < x&) Then
        x& = (ScaleWidth - .Width) - 90
    End If
    .Left = x&
End With

If (para.Newline) Then
    If (para.MARS) Then
        With lblMarsModus
            .Left = .Left + 2 * wpara.NlFlexBackY
            .Top = .Top + 2 * wpara.NlFlexBackY
        End With
    End If
    With lblRezeptNr
        .Left = .Left + 2 * wpara.NlFlexBackY
        .Top = .Top + 2 * wpara.NlFlexBackY
    End With
    With txtRezeptNr
        .Left = lblRezeptNr.Left
        .Top = .Top + 2 * wpara.NlFlexBackY + 90
        .Width = TextWidth(String(25, "9")) + 90
        
        .BackColor = vbWhite
        Call wpara.ControlBorderless(txtRezeptNr, 2, 2)
    End With
    If (para.MARS) Then
        With lblMarsModus
            .Width = txtRezeptNr.Width + 120
        End With
    End If

    
    With flxKkassen
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .GridLines = flexGridNone

        If (.Rows = 0) Then .Rows = 1
        .Left = lblRezeptNr.Left
        .Top = txtRezeptNr.Top + txtRezeptNr.Height + 450
        
        .Width = flxarbeit(0).Left + 2 * txtRezeptNr.Width - 2 * .Left
        .Width = txtRezeptNr.Width - 300
        
'        With flxInfo(0)
'            x = .Left + .Width - 75
'        End With
'        x = x - picRezept.Width - (.Left + 1200)
'        If (x < .Width) Then
'            .Width = x
'        End If
        
        .Height = opBereich.ArbeitBackHeight - .Top - 300
        .Height = (.Height \ .RowHeight(0)) * .RowHeight(0) + 90

        .ColAlignment(0) = flexAlignLeftCenter
        .ColWidth(0) = 0
        .ColWidth(1) = .Width
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .BackColor = RGB(232, 217, 172)
        .BackColorBkg = RGB(232, 217, 172)

        x = .Left + .Width + 600
        x = txtRezeptNr.Left + txtRezeptNr.Width + 300
    End With
    
    With flxVerordnung
        .Rows = 2
        .FixedRows = 1
        .FormatString = ">Preis|>Menge|Meh|<Kurzbezeichnung|<PZN||Flag|KP|Gstufe|PE_PM|PE_AI|PE_AnzEinheiten|Verwurf|>PreisUNgerundet"
        .Rows = 1
        
        .ColWidth(0) = TextWidth("9999999.99")
        .ColWidth(1) = TextWidth(String(8, "9"))
        .ColWidth(2) = TextWidth(String(6, "A"))
        .ColWidth(3) = TextWidth(String(35, "A"))
        .ColWidth(4) = TextWidth(String(9, "9"))
        .ColWidth(5) = wpara.FrmScrollHeight
        For i = 6 To 13
            .ColWidth(i) = 0
        Next i
        
        Dim Breite1%
        Breite1% = 0
        For i% = 0 To (.Cols - 1)
            Breite1% = Breite1% + .ColWidth(i%)
        Next i%
        .Width = Breite1% + 90
        
        Dim iArbeitAnzzeilen%
        iArbeitAnzzeilen% = 10
        .Height = .RowHeight(0) * iArbeitAnzzeilen% + 90
        
    '    If (MagSpeicherIndex% > 0) Then
    '        Call HoleMagSpeicher
    '    End If
    '
    '    .AddItem " "
    '    .row = .Rows - 1
    End With
    
    picBack(0).FillColor = RGB(232, 217, 172)
    With flxarbeit(0)
        RoundRect picBack(0).hdc, (.Left) / Screen.TwipsPerPixelX, (.Top / Screen.TwipsPerPixelY) - 5, (.Left + x) / Screen.TwipsPerPixelX, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
        Call wpara.FillGradient(picBack(0), (.Left) / Screen.TwipsPerPixelX + 1, ((.Top + .Height / 2) / Screen.TwipsPerPixelY), (.Left + x) / Screen.TwipsPerPixelX - 2, (.Top + .Height) / Screen.TwipsPerPixelY - 60, RGB(232, 217, 172), RGB(242, 227, 182))
    End With
    
    With picBack(0)
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtRezeptNr.Left - 60) / Screen.TwipsPerPixelX, (txtRezeptNr.Top - 30) / Screen.TwipsPerPixelY, (txtRezeptNr.Left + txtRezeptNr.Width + 60) / Screen.TwipsPerPixelX, (txtRezeptNr.Top + txtRezeptNr.Height + 15) / Screen.TwipsPerPixelY, 10, 10
    
        .Refresh
    End With

'    With flxKkassen
'        .ScrollBars = flexScrollBarNone
'        .BorderStyle = 0
'        .GridLines = flexGridNone
'
'        If (.Rows = 0) Then .Rows = 1
'        .Left = lblRezeptNr.Left
'        .Top = txtRezeptNr.Top + txtRezeptNr.Height + 450
'
'        .Width = flxarbeit(0).Left + 2 * txtRezeptNr.Width - 2 * .Left
'        x = ScaleWidth - picRezept.Width - (.Left + 1200)
'        If (x < .Width) Then
'            .Width = x
'        End If
'
'        .Height = opBereich.ArbeitBackHeight - .Top - 300
'        .Height = (.Height \ .RowHeight(0)) * .RowHeight(0) + 90
'
'        .ColAlignment(0) = flexAlignLeftCenter
'        .ColWidth(0) = 0
'        .ColWidth(1) = .Width
'        .ColWidth(2) = 0
'        .ColWidth(3) = 0
'        .BackColor = RGB(232, 217, 172)
'        .BackColorBkg = RGB(232, 217, 172)
''        .GridColor = wpara.FarbeArbeit
'    End With
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

On Error Resume Next
lblRezeptNr.BackColor = wpara.FarbeArbeit
fmeAv.BackColor = wpara.FarbeArbeit

For i% = 0 To 2
    lblAv(i%).BackColor = wpara.FarbeArbeit
Next i%

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
Dim i%, j%, k%, ind%, iVal%, erg%
Dim l&, f&, lVal&, lColor&
Dim h$, h2$, Key$, wert1$, BetrLief$, Lief2$
Dim sPzn$, sKkBez$, sKassenId$, sStatus$, sGültigBis$
    
With Me
    
'    .Height = wpara.WorkAreaHeight
'    If (wpara.BildFaktor = 0.8!) Then
'        .Width = 10980 * wpara.BildFaktor '10200
'    Else
'        .Width = .Width * wpara.BildFaktor
'    End If
'
'
'    iVal% = (Screen.Width - Width) / 2
'    If (iVal% < 0) Then
'        iVal% = 0
'    End If
'    h$ = Format(iVal%, "00000")
'    l& = GetPrivateProfileString(UserSection$, "StartX", h$, h$, 6, INI_DATEI)
'    h$ = Left$(h$, l&)
'    lVal& = Val(h$)
'    If (lVal& > 30000) Then
'        iVal% = -9999
'    Else
'        iVal% = lVal&
'    End If
'    If (iVal% = -9999) Then
'        WindowState = vbMaximized
'    Else
'        If (iVal% < 0) Then
'            iVal% = 0
'        End If
'        .Left = iVal%
'
'        iVal% = 75
'        h$ = Format(iVal%, "00000")
'        l& = GetPrivateProfileString(UserSection$, "StartY", h$, h$, 6, INI_DATEI)
'        h$ = Left$(h$, l&)
'        iVal% = Val(h$)
'        If (iVal% < 0) Then
'            iVal% = 0
'        End If
'        .Top = iVal%
'
'        iVal% = .Width
'        h$ = Format(iVal%, "00000")
'        l& = GetPrivateProfileString(UserSection$, "BreiteX", h$, h$, 6, INI_DATEI)
'        h$ = Left$(h$, l&)
'        iVal% = Val(h$)
'        If (iVal% < 0) Then
'            iVal% = 0
'        End If
'        If (.Left + iVal% > wpara.WorkAreaWidth) Then
'            WindowState = vbMaximized
'        Else
'            .Width = iVal%
'
'            iVal% = .Height
'            h$ = Format(iVal%, "00000")
'            l& = GetPrivateProfileString(UserSection$, "HoeheY", h$, h$, 6, INI_DATEI)
'            h$ = Left$(h$, l&)
'            iVal% = Val(h$)
'            If (iVal% < 0) Then
'                iVal% = 0
'            End If
'            If (.Top + iVal% > wpara.WorkAreaHeight) Then
'                WindowState = vbMaximized
'            Else
'                .Height = iVal%
'            End If
'        End If
'    End If
    Call wpara.InitForm(Me, UserSection$, INI_DATEI)
    
    
    ProgrammTyp% = 0
    If (Command <> "") Then ProgrammTyp% = Val(Command)
    



    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerVersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptVersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptVersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "PrivatRezeptDruckerVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    PrivatRezeptVersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDatumVersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    DatumVersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptDatumVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    DatumVersatzY% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "RezeptNrVersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptNrVersatzY% = Val(h$)
            
'    h2$ = Left$("Lucida Console" + Space$(20), 20)
    h2$ = Left$("Courier New" + Space$(20), 20)
    h$ = Space$(20)
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerSchriftArt", h2$, h$, 21, INI_DATEI)
    h$ = Left$(h$, l&)
    RezeptFont$ = RTrim$(h$)
    
    
''' Code 128
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "Code128VersatzX", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    Code128VersatzX% = Val(h$)
            
    iVal% = 0
    h$ = Format(iVal%, "00000")
    l& = GetPrivateProfileString(UserSection$, "Code128VersatzY", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    Code128VersatzY% = Val(h$)
'    If (RezeptVersatzY% < 0) Then RezeptVersatzY% = 0
            
    h2$ = Left$("Arial,8" + Space$(20), 20)
    h$ = Space$(20)
    l& = GetPrivateProfileString(UserSection$, "Code128SchriftArt", h2$, h$, 21, INI_DATEI)
    h$ = Left$(h$, l&)
    ind% = InStr(h$, ",")
    Code128Font$ = RTrim$(Left$(h$, ind% - 1))
    Code128FontSize = Val(Mid$(h$, ind% + 1))
'''''''''''''''''''''''''''
    
    
    
    For i% = 0 To 3
        If (i% = 0) Then
            h$ = "DCD2FA"
        ElseIf (i% = 2) Then
            h$ = "BEFFFF"
        ElseIf (i% = 3) Then
            h$ = "FAE4CF"
        Else
            h$ = "DCD2FA"
        End If
        
        Key$ = "FarbeRezept" + Format(i%, "0")
        l& = GetPrivateProfileString("Rezeptkontrolle", Key$, h$, h$, 21, INI_DATEI)
        h$ = Left$(h$, l&)
        
        RezeptFarben&(i%) = wpara.BerechneFarbWert(h$)
    Next i%



    For i% = 0 To 6
        h$ = Space$(20)
        Key$ = "Darstellung" + Format(i%, "0")
        l& = GetPrivateProfileString("Taxierung", Key$, h$, h$, 21, INI_DATEI)
        h$ = Left$(h$, l&)
        
        If (para.Newline) Then
            MagDarstellung&(i%, 0) = vbBlack
            MagDarstellung&(i%, 1) = vbWhite
        Else
            MagDarstellung&(i%, 0) = flxInfo(0).ForeColor
            MagDarstellung&(i%, 1) = flxInfo(0).BackColor
        End If
    
        If (Trim(h$) <> "") Then
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
'                lColor& = wpara.BerechneFarbWert(Mid$(h$, ind% + 1))
'                If (lColor& <> 0) Then MagDarstellung&(i%, 1) = lColor&
                MagDarstellung&(i%, 1) = wpara.BerechneFarbWert(Mid$(h$, ind% + 1))
                h$ = Left$(h$, ind% - 1)
            End If
'            lColor& = wpara.BerechneFarbWert(h$)
'            If (lColor& <> 0) Then MagDarstellung&(i%, 0) = lColor&
            If (Trim(h$) <> "") Then
                MagDarstellung&(i%, 0) = wpara.BerechneFarbWert(h$)
            End If
        End If
    Next i%

    h$ = Space$(50)
    l& = GetPrivateProfileString(UserSection$, "RezeptDrucker", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    RezeptDrucker$ = h$
    
    h$ = Space$(50)
    l& = GetPrivateProfileString(UserSection$, "RezeptDruckerParameter", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    RezeptDruckerPara$ = h$



'    i% = 1
    j% = 0
    For i = 1 To 2
        h$ = Space$(100)
        Key$ = "Taetigkeit" + Format(i%, "00")
        l& = GetPrivateProfileString("Rezeptkontrolle", Key$, " ", h$, 101, INI_DATEI)
        h$ = Trim$(h$)
        If (Len(h$) > 1) Then
            h$ = Left$(h$, Len(h$) - 1)
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                wert1$ = RTrim$(Left$(h$, ind% - 1))
                BetrLief$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
                If (wert1$ <> "") Then
                    Taetigkeiten(j%).Taetigkeit = wert1$
                    
                    For k% = 0 To 79
                        ind% = InStr(BetrLief$, ",")
                        If (ind% > 0) Then
                            Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                            BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                        Else
                            Lief2$ = BetrLief$
                            BetrLief$ = ""
                        End If
                        If (Lief2$ <> "") Then
                            Taetigkeiten(j%).pers(k%) = Val(Lief2$)
                        End If
                        If (BetrLief$ = "") Then Exit For
                    Next k%
                    j% = j% + 1
                End If
            End If
        End If
    Next i
    AnzTaetigkeiten% = j%

    
    j% = 0
    For i% = 1 To 10
        h$ = Space$(100)
        Key$ = "SonderBeleg" + Format(i%, "00")
        l& = GetPrivateProfileString("Rezeptkontrolle", Key$, " ", h$, 101, INI_DATEI)
        h$ = Trim$(h$)
        If (Len(h$) > 1) Then
            h$ = Left$(h$, Len(h$) - 1)
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sPzn$ = RTrim$(Left$(h$, ind% - 1))
                If (Len(sPzn) = 7) Then
                    sPzn = "0" + sPzn
                End If
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sKkBez$ = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sKassenId$ = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sStatus = RTrim$(Left$(h$, ind% - 1))
                h$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            End If
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                sGültigBis$ = RTrim$(Left$(h$, ind% - 1))
                With SonderBelege(j%)
                    .pzn = sPzn$
                    .KkBez = sKkBez$
                    .KassenId = sKassenId$
                    .Status = sStatus$
                    .GültigBis = sGültigBis$
                    j% = j% + 1
                End With
            End If
        End If
    Next i%
    AnzSonderBelege% = j%

    
    h$ = "J"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezepturMitFaktor", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        RezepturMitFaktor% = True
    Else
        RezepturMitFaktor% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "BtmAlsZeile", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        BtmAlsZeile% = True
    Else
        BtmAlsZeile% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezepturDruck", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        RezepturDruck% = True
    Else
        RezepturDruck% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "AvpTeilnahme", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        AvpTeilnahme% = True
    Else
        AvpTeilnahme% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "Code128Aktiv", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        Code128Flag% = True
    Else
        Code128Flag% = False
    End If

    
    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "InstitutsKz", h$, h$, 51, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    If (h$ <> "") Then h$ = Right$(Space$(7) + Trim(Left$(h$, l&)), 7)
    RezApoNr$ = h$
    OrgRezApoNr$ = RezApoNr$
    
    h$ = "30"
    l& = GetPrivateProfileString("Rezeptkontrolle", "InstitutsKzPraefix", h$, h$, 3, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    iVal = Val(h)
    If (iVal <= 0) Or (iVal >= 100) Then
        iVal = 30
    End If
    RezApoNrPraefix$ = Format(iVal, "00")

    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezeptText", h$, h$, 51, INI_DATEI)
    RezApoDruckName$ = Trim(Left$(h$, l&))
    
    h$ = Space$(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "BtmRezeptText", h$, h$, 51, INI_DATEI)
    BtmRezDruckName$ = Trim(Left$(h$, l&))
    If (BtmRezDruckName$ = "") Then
        BtmRezDruckName$ = RezApoDruckName$
    End If
    
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "ArtIndexDebug", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        ArtIndexDebug% = True
    Else
        ArtIndexDebug% = False
    End If

    For i% = 0 To 15    'UBound(ParenteralPzn)
        h$ = Left$(ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00") + Space$(20), 20)
        Key$ = "SonderPzn" + CStr(i)
        l& = GetPrivateProfileString("Parenteral", Key$, h$, h$, 21, INI_DATEI)
        h$ = Left$(h$, l&)
        ind = InStr(h, ";")
        If (ind > 0) Then
            ParenteralPzn(i) = Left$(h, ind - 1)
            ParenteralPreis(i) = xVal(Mid$(h, ind + 1))
        End If
    Next i%
    i = 6
    If (ParenteralPzn(i) = "9999092") Then
        ParenteralPzn$(i) = "2567478"
        h$ = ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i), h$, INI_DATEI)
        
        ParenteralPzn$(i + 8) = "2567478"
        h$ = ParenteralPzn(i + 8) + ";" + Format(ParenteralPreis(i + 8), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i + 8), h$, INI_DATEI)
    End If
    i = 7
    If (ParenteralPzn(i) = "9999152") Then
        ParenteralPzn$(i) = "2567461"
        h$ = ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i), h$, INI_DATEI)
        
        ParenteralPzn$(i + 8) = "2567461"
        h$ = ParenteralPzn(i + 8) + ";" + Format(ParenteralPreis(i + 8), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i + 8), h$, INI_DATEI)
    End If
    For i% = 0 To UBound(ParenteralPzn)
        If (Len(ParenteralPzn(i)) = 7) Then
            ParenteralPzn(i) = "0" + ParenteralPzn(i)
        End If
        h$ = ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i), h$, INI_DATEI)
    Next i%


    For i% = 0 To UBound(ParEnteralAufschlag)
        h$ = Left$(Format(ParEnteralAufschlag(i), "0.00") + Space$(20), 20)
        Key$ = "Aufschlag" + CStr(i)
        l& = GetPrivateProfileString("Parenteral", Key$, h$, h$, 21, INI_DATEI)
        h$ = Left$(h$, l&)
        ParEnteralAufschlag(i) = xVal(h)
    Next i%

    For i = 1 To 4
        h$ = Space$(50)
        l& = GetPrivateProfileString("Parenteral", "HerstellerKz" + CStr(i), h$, h$, 51, INI_DATEI)
        h$ = Trim(Left$(h$, l&))
        h$ = Right$(String(9, "0") + CStr(Val(h$)), 9)
        ParEnteralHerstellerKz$(i) = h$
    Next i
    
    TmCheck = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "TmCheck", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        TmCheck% = True
    End If

    BenutzerSignatur = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "Signatur", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        BenutzerSignatur% = True
    End If

    NurHashCodeDruck = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "HashCodeDruck", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        NurHashCodeDruck% = True
    End If

    RezeptNrPositionAlt = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "RezeptNrPositionAlt", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        RezeptNrPositionAlt = True
    End If

    Pzn8Test = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "Pzn8Test", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        Pzn8Test = True
    End If

    PreisDiffAktiv = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "PreisDiff", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        PreisDiffAktiv = True
    End If

    h$ = Space$(2000)
    l& = GetPrivateProfileString("Rezeptkontrolle", "Beigetreten", h$, h$, 2001, INI_DATEI)
    h$ = Trim(Left$(h$, l&))
    If (h$ <> "") Then h$ = "," + h$ + ","
    BeigetreteneVereinbarungen$ = h$
    
    h$ = "J"
    l& = GetPrivateProfileString("Rezeptkontrolle", "FormatPrivatRezept", "J", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        FormatPrivatRezept = True
    Else
        FormatPrivatRezept = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "DatumObenPrivatRezept", "N", h$, 2, INI_DATEI)
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        DatumObenPrivatRezept = True
    Else
        DatumObenPrivatRezept = False
    End If


'    h$ = "01"
'    l& = GetPrivateProfileString(INI_SECTION, "MinutenVerspaetung", "3", h$, 3, INI_DATEI)
'    h$ = Left$(h$, l&)
'    AnzMinutenVerspaetung% = Val(h$)
'
'    h$ = "N"
'    l& = GetPrivateProfileString(INI_SECTION, "BestVorsKomplett", "N", h$, 2, INI_DATEI)
'    h$ = Left$(h$, l&)
'    If (h$ = "J") Then
'        BestVorsKomplett% = True
'    Else
'        BestVorsKomplett% = False
'    End If
'
'    h$ = Space$(8)
'    l& = GetPrivateProfileString(UserSection$, "FarbeGray", h$, h$, 9, INI_DATEI)
'    h$ = Left$(h$, l&)
'    If (Trim(h$) = "") Then
'        FarbeGray& = vbGrayText
'    Else
'        FarbeGray& = wpara.BerechneFarbWert(h$)
'    End If
'
'    h$ = "100"
'    l& = GetPrivateProfileString(INI_SECTION, "SchwellwertWarnungProz", "100", h$, 4, INI_DATEI)
'    h$ = Left$(h$, l&)
'    SchwellwertWarnungProz% = Val(h$)

    h$ = "N"
    l& = GetPrivateProfileString("Allgemein", "Debug", "N", h$, 2, CurDir + "\winrezk.ini")
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        WinRezDebugAktiv% = True
    Else
        WinRezDebugAktiv% = False
    End If

    h$ = "N"
    l& = GetPrivateProfileString("Allgemein", "PreisKz_62_70", "N", h$, 2, CurDir + "\FiveRx.ini")
    h$ = Left$(h$, l&)
    If (h$ = "J") Then
        PreisKz_62_70 = True
    Else
        PreisKz_62_70 = False
    End If

    AlleRezepturenMitHash = False
    h$ = "N"
    l& = GetPrivateProfileString("Rezeptkontrolle", "AlleRezepturenMitHash", "N", h$, 2, INI_DATEI)
    h$ = UCase(Left$(h$, l&))
    If (h$ = "J") Then
        AlleRezepturenMitHash = True
    End If

    LennartzPfad = ""
    h$ = Space(50)
    l& = GetPrivateProfileString("Rezeptkontrolle", "LennartzPfad", h, h$, 51, INI_DATEI)
    LennartzPfad = Trim(Left$(h$, l&))
    
    ChefModus = False
    h = "0"
    l = GetPrivateProfileString("Allgemein", "ChefModus", h$, h$, 2, CurDir() + "\Ti_back.ini")
    If (l > 0) Then
        ChefModus = (Val(Trim(Left(h$, l))) = 1)
    End If
End With

If (para.Newline) Then
    RezeptFarben&(0) = RGB(255, 170, 120)   ' RGB(255, 131, 62)
End If

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

For i% = 1 To 13
    Load cmdDatei(i%)
Next i%

For i% = 0 To 10
    cmdDatei(i%).Top = 0
    cmdDatei(i%).Left = i% * 900
    cmdDatei(i%).Visible = True
    cmdDatei(i%).ZOrder 1
Next i%

cmdDatei(0).Caption = "&S"
cmdDatei(1).Caption = "&I"
cmdDatei(2).Caption = "&H"
cmdDatei(3).Caption = "&P"
cmdDatei(4).Caption = "&F"
cmdDatei(5).Caption = "&V"
cmdDatei(6).Caption = "&A"
cmdDatei(7).Caption = "&R"
cmdDatei(8).Caption = "&Z"
cmdDatei(9).Caption = "&E"
cmdDatei(10).Caption = "&M"
cmdDatei(11).Caption = "&L"
cmdDatei(12).Caption = "&T"
cmdDatei(13).Caption = "&."

For i% = 1 To 2
    Load cboVerfügbarkeit(i%)
Next i%

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
Dim x&, y&, nHWnd&, l&
Dim h$
             
tmrAction.Enabled = False

If (WinVkId > 0) Then
    If (CheckTask%(WinVkId&)) Then
        tmrAction.Interval = 2000
        tmrAction.Enabled = True
    Else
        WinVkId = 0
    
        h$ = "0000001"
        l& = GetPrivateProfileString("Rezeptkontrolle", "ANSGVerkauf", h$, h$, 8, INI_DATEI)
        h$ = Left$(h$, l&)
        
'        With picRezept
'            .FontSize = .FontSize + 2
'            .FontBold = False
'            .CurrentX = WinVkX
'            .CurrentY = WinVkY
'            picRezept.Print WinVkText + h
'            .FontSize = .FontSize - 2
'        End With
        With lblWinVk
            .ForeColor = vbBlack
            .Caption = WinVkText + h
            .Visible = True
            Call ActProgram.PaintSelbsterklaerung(h)
        End With
        
    End If
Else
    ' Handle der MsgBox ermitteln
    DoEvents
    nHWnd = GetActiveWindow()
    
    ' MsgBox positionieren
    x = (picRezept.Left + KundenBoxX) / Screen.TwipsPerPixelX + 10
    y = (wpara.FrmCaptionHeight + wpara.FrmMenuHeight + picBack(0).Top + picRezept.Top + KundenBoxY) / Screen.TwipsPerPixelY + 10
    SetWindowPos nHWnd, Me.hWnd, x, y, 0, 0, 1 'SWP_NOSIZE
End If

Call DefErrPop
End Sub

Private Sub tmrF6Sperre_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrF6Sperre_Timer")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
             
tmrF6Sperre.Enabled = False

Call DefErrPop
End Sub

Private Sub txtRezeptNr_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtRezeptNr_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

With txtRezeptNr
    .SelStart = 0
    .SelLength = Len(.text)
End With

If (Chef) Then
    mnuBearbeitenInd(MENU_SF3).Enabled = True
    mnuBearbeitenInd(MENU_SF4).Enabled = VmFlag%
    mnuBearbeitenInd(MENU_SF6).Enabled = True
    mnuBearbeitenInd(MENU_SF7).Enabled = True
    mnuBearbeitenInd(MENU_SF8).Enabled = True
    mnuBearbeitenInd(MENU_SF9).Enabled = VmFlag%
    cmdToolbar(MENU_SF3).Enabled = True
    cmdToolbar(MENU_SF4).Enabled = VmFlag%
    cmdToolbar(MENU_SF6).Enabled = True
    cmdToolbar(MENU_SF7).Enabled = True
    cmdToolbar(MENU_SF8).Enabled = True
    cmdToolbar(MENU_SF9).Enabled = VmFlag%
Else
    MarsNoChef
End If

For i% = 0 To 8
    mnuBearbeitenZusatz(i%).Enabled = False
Next i%
'mnuLennartz.Enabled = False

EingabeStr$ = ""

Call DefErrPop
End Sub

Private Sub ResetToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResetToolbar")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

h = "1717"
'    Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, "1717")  'weil in DLL gelöscht !
If (para.Newline) Then
    Call opToolbar.InitToolbar(Me, App.EXEName, INI_SECTION, h)
Else
    Call opToolbar.InitToolbar(Me, INI_DATEI, INI_SECTION, h)
End If
Call ShowToolbar

cmdToolbar(5).ToolTipText = "F6 Drucken des Rezepts"
cmdToolbar(9).ToolTipText = "shift+F2 Impfstoff-Rezept"
cmdToolbar(13).ToolTipText = "shift+F6 Neues Rezept - zuz.frei"
cmdToolbar(14).ToolTipText = "shift+F7 Neues Rezept - zuz.pflichtig"

Call DefErrPop
End Sub

Public Sub cmdAv_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAv_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

frmAvAuswahl.Show 1
Call ActProgram.PaintAvParameter

If (picRezept.Visible) Then
    picRezept.SetFocus
Else
    txtRezeptNr.SetFocus
End If

Call DefErrPop
End Sub

Private Sub flxKkassen_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKKassen_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxKkassen
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxKkassen_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKKassen_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxKkassen
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub flxKkassen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKKassen_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call ActProgram.flxKkassenClick

Call DefErrPop
End Sub

Sub ErzeugeDruckerAuswahl()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErzeugeDruckerAuswahl")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, OpDrucker_Log%
Dim h$

IstDosDrucker% = False

On Error Resume Next
For i% = 3 To 19
    Unload mnuDruckerInd(i%)
Next i%
For i% = 2 To 19
    Unload MnuDosDruckerInd(i%)
Next i%
On Error GoTo DefErr

For i% = 2 To 5
    Load MnuDosDruckerInd(i%)
Next i%
MnuDosDruckerInd(2).Caption = "TM290"
MnuDosDruckerInd(3).Caption = "TM290-II"
MnuDosDruckerInd(4).Caption = "TM-U950"
MnuDosDruckerInd(5).Caption = "TM5000II"

For i% = 2 To 5
    h$ = MnuDosDruckerInd(i%).Caption
    If (h$ = RezeptDrucker$) Then
        MnuDosDruckerInd(i%).Checked = True
        IstDosDrucker% = True
'        mnuDruckerInd(0).Checked = True
'        Set Printer = Printers(i%)
    Else
        MnuDosDruckerInd(i%).Checked = False
    End If
    MnuDosDruckerInd(i%).Enabled = True
Next i%

MnuDosDruckerInd(0).Enabled = IstDosDrucker%
    

'For i% = 0 To (Printers.Count - 1)
'    j% = i% + 3
'    Load mnuDruckerInd(j%)
'    h$ = Printers(i%).DeviceName
'    mnuDruckerInd(j%).Caption = h$
'    If (h$ = RezeptDrucker$) Then
'        mnuDruckerInd(j%).Checked = True
'        Set Printer = Printers(i%)
'    Else
'        mnuDruckerInd(j%).Checked = False
'    End If
'Next i%

    OpDrucker_Log% = FreeFile
    Open "OpDrucker.log" For Append As #OpDrucker_Log
    Print #OpDrucker_Log%, Format(Now, "DD.mm.YYYY HH:mm:SS ") + "REZEPTDRUCKER: " + " <" + RezeptDrucker + ">"
    Close #OpDrucker_Log%
    
Dim hPrinter As Printer
Call wpara.InstalledPrinters
For i = 0 To (wpara.PrinterCount - 1)
    j% = i% + 3
    Load mnuDruckerInd(j%)
    h$ = wpara.PrinterName(i + 1)
    h = wpara.PrinterNameOP(h)
    mnuDruckerInd(j%).Caption = h$
            
    OpDrucker_Log% = FreeFile
    Open "OpDrucker.log" For Append As #OpDrucker_Log
    Print #OpDrucker_Log%, Format(Now, "DD.mm.YYYY HH:mm:SS ") + "INSTALLIERT: " + " <" + wpara.PrinterName(i + 1) + ">" + "  <" + h + ">"
    Close #OpDrucker_Log%
    
    If (UCase(h$) = UCase(RezeptDrucker$)) Then
        mnuDruckerInd(j%).Checked = True
'        For k% = 0 To (Printers.Count - 1)
        For Each hPrinter In Printers
            OpDrucker_Log% = FreeFile
            Open "OpDrucker.log" For Append As #OpDrucker_Log
            Print #OpDrucker_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + "VB-PRINTERS: " + "<" + hPrinter.DeviceName + ">"
            Close #OpDrucker_Log%
            
            If (UCase(wpara.PrinterName(i + 1)) = UCase(hPrinter.DeviceName)) Then
                OpDrucker_Log% = FreeFile
                Open "OpDrucker.log" For Append As #OpDrucker_Log
                Print #OpDrucker_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + "VB-PRINTERS: " + "<" + hPrinter.DeviceName + "> GEFUNDEN"
                Close #OpDrucker_Log%
                
                Set Printer = hPrinter
                Exit For
            End If
        Next
    Else
        mnuDruckerInd(j%).Checked = False
    End If
Next i%
    

mnuDruckerInd(2).Enabled = (IstDosDrucker% = False)

mnuCode128.Enabled = Code128Flag


Call DefErrPop
End Sub

Private Sub mnuDruckerInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDruckerInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

If (index > 2) Then
    RezeptDrucker$ = mnuDruckerInd(index).Caption
    
    l& = WritePrivateProfileString(UserSection$, "RezeptDrucker", RezeptDrucker$, INI_DATEI)
    
    'Set Printer = Printers(Index)
    Call ErzeugeDruckerAuswahl
End If

Call DefErrPop
End Sub

Private Sub mnuDosDruckerInd_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDosDruckerInd_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

If (index = 0) Then
    h$ = MyInputBox("Schnittstellen-Parameter: ", "Drucker ohne Windows-Treiber", RezeptDruckerPara$)
    h$ = UCase(Trim(h$))
    If (h$ <> "") Then
        RezeptDruckerPara$ = h$
        l& = WritePrivateProfileString(UserSection$, "RezeptDruckerParameter", RezeptDruckerPara$, INI_DATEI)
    End If
Else
    RezeptDrucker$ = MnuDosDruckerInd(index).Caption
    l& = WritePrivateProfileString(UserSection$, "RezeptDrucker", RezeptDrucker$, INI_DATEI)
    
    If (Left$(RezeptDrucker$, 5) = "TM290") Then
        RezeptDruckerPara$ = "COM2:4800,N,8,1,RS"
    Else
        RezeptDruckerPara$ = "COM2:9600,N,8,1,RS"
    End If
    l& = WritePrivateProfileString(UserSection$, "RezeptDruckerParameter", RezeptDruckerPara$, INI_DATEI)
    
    Call ErzeugeDruckerAuswahl
End If

Call DefErrPop
End Sub

Private Sub mnuDruckerWinPara_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuDruckerWinPara_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim l&
Dim h$

If (index = 0) Then
    h$ = MyInputBox("Druck-Versatz in X-Richtung: ", "Drucker mit Windows-Treiber", Str$(RezeptVersatzX%))
    h$ = UCase(Trim(h$))
    If (h$ <> "") And (Val(h$) >= 0) Then
        RezeptVersatzX% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "RezeptDruckerVersatzX", Str$(RezeptVersatzX%), INI_DATEI)
    End If
ElseIf (index = 1) Then
    h$ = MyInputBox("Druck-Versatz in Y-Richtung: ", "Drucker mit Windows-Treiber", Str$(RezeptVersatzY%))
    h$ = UCase(Trim(h$))
    If (h$ <> "") And (Val(h$) >= 0) Then
        RezeptVersatzY% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "RezeptDruckerVersatzY", Str$(RezeptVersatzY%), INI_DATEI)
    End If
ElseIf (index = 2) Then
    h$ = MyInputBox("Druck-Versatz des Datums in X-Richtung: ", "Drucker mit Windows-Treiber", Str$(DatumVersatzX%))
    h$ = UCase(Trim(h$))
'    If (h$ <> "") And (Val(h$) >= 0) Then
    If (h$ <> "") Then
        DatumVersatzX% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "RezeptDatumVersatzX", Str$(DatumVersatzX%), INI_DATEI)
    End If
ElseIf (index = 3) Then
    h$ = MyInputBox("Druck-Versatz des Datums in Y-Richtung: ", "Drucker mit Windows-Treiber", Str$(DatumVersatzY%))
    h$ = UCase(Trim(h$))
'    If (h$ <> "") And (Val(h$) >= 0) Then
    If (h$ <> "") Then
        DatumVersatzY% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "RezeptDatumVersatzY", Str$(DatumVersatzY%), INI_DATEI)
    End If
ElseIf (index = 4) Then
    h$ = MyInputBox("Druck-Versatz der RezeptNummer in Y-Richtung: ", "Drucker mit Windows-Treiber", Str$(RezeptNrVersatzY%))
    h$ = UCase(Trim(h$))
    If (h$ <> "") And (Val(h$) >= 0) Then
        RezeptNrVersatzY% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "RezeptNrVersatzY", Str$(RezeptNrVersatzY%), INI_DATEI)
    End If
Else
    erg% = EditFont%
    If (erg%) Then l& = WritePrivateProfileString(UserSection$, "RezeptDruckerSchriftArt", RezeptFont$, INI_DATEI)
End If

Call DefErrPop
End Sub

Private Sub mnuCode128Ind_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuCode128Ind_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim l&
Dim h$

If (index = 0) Then
    h$ = MyInputBox("Druck-Versatz in X-Richtung: ", "Code 128", Str$(Code128VersatzX%))
    h$ = UCase(Trim(h$))
    If (h$ <> "") Then  'And (Val(h$) >= 0) Then
        Code128VersatzX% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "Code128VersatzX", Str$(Code128VersatzX%), INI_DATEI)
    End If
ElseIf (index = 1) Then
    h$ = MyInputBox("Druck-Versatz in Y-Richtung: ", "Code 128", Str$(Code128VersatzY%))
    h$ = UCase(Trim(h$))
    If (h$ <> "") Then  'And (Val(h$) >= 0) Then
        Code128VersatzY% = Val(h$)
        l& = WritePrivateProfileString(UserSection$, "Code128VersatzY", Str$(Code128VersatzY%), INI_DATEI)
    End If
Else
    erg% = EditCode128Font%
    If (erg%) Then l& = WritePrivateProfileString(UserSection$, "Code128SchriftArt", RTrim$(Code128Font) + "," + CStr(Code128FontSize), INI_DATEI)
End If

Call DefErrPop
End Sub

Public Function ArbeitBackHeight%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ArbeitBackHeight%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
ArbeitBackHeight% = opBereich.ArbeitBackHeight
Call DefErrPop
End Function

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
'        OrgY = Y

'        Call opToolbar.ShowToolbar
        Call opToolbar.MouseMove(x)
    End If
End If

Call DefErrPop
End Sub

Sub ShowToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowToolbar")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    Call opToolbar.ShowToolbar
End If

Call DefErrPop
End Sub

Sub MarsNoChef()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MarsNoChef")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

If (Chef = False) Then
    For i = 0 To 8
        mnuDateiInd(i).Enabled = False
    Next i
    For i = 0 To 7
        mnuBearbeitenInd(i).Enabled = False
    Next i
    For i = 8 To 15
        mnuBearbeitenInd(i + 1).Enabled = False
    Next i
    
    mnuBearbeitenLayout.Enabled = False
    
    On Error Resume Next
    For i% = 0 To 8
        mnuBearbeitenZusatz(i%).Enabled = False
    Next i%
    mnuLennartz.Enabled = False
    On Error GoTo DefErr
    
    mnuPrivRezVK.Enabled = False
    
    'mnuBearbeiten.Enabled = False
    mnuAnsicht.Enabled = False
    mnuExtras.Enabled = False
    
    For i = 0 To 2
        cboVerfügbarkeit(i).Enabled = False
    Next
End If

Call DefErrPop
End Sub

Sub MarsInitAutomatikModus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MarsInitAutomatikModus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim iDat As Date

lstMarsRezepte.Clear

iDat = DateValue("01.07.2016")
h = Space(255)
l = GetPrivateProfileString("Rezeptkontrolle", "LetztDatum", "", h, 255, CurDir + "\Ocr2016.ini")
If l > 0 Then
    h = Left(h, l)
    If IsDate(h) Then
        iDat = h
    End If
End If

SQLStr = "SELECT * FROM Verkauf"
SQLStr = SQLStr + " WHERE ((RezeptNr<>'') OR (PrivRezLaufNr>0))"
'SQLStr = SQLStr + " AND (Datum<='" + Left(h, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5, 2) + " 23:59" + "')"
SQLStr = SQLStr + " AND (Datum>='" + Format(iDat, "DD.MM.YYYY HH:mm:ss") + "')"
SQLStr = SQLStr + " ORDER BY Datum,LaufNr,ZeilenNr"
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
If (FabsErrf <> 0) Then
    Call iMsgBox("keine passenden Rezepte gespeichert !")
    Call DefErrPop: Exit Sub
End If
VerkaufAdoRec.MoveFirst
Do
    If (VerkaufAdoRec.EOF) Then
        Exit Do
    End If
    
    h = Trim(CheckNullStr(VerkaufAdoRec!RezeptNr))
    If (h = "") Then
        h = CStr(CheckNullLong(VerkaufAdoRec!PrivRezlaufNr))
    End If
    
    
    With lstMarsRezepte
        For i = 0 To (.ListCount - 1)
            If (.List(i) = h) Then
                .RemoveItem (i)
                Exit For
            End If
        Next i
        
        If (TesteMarsRezept(h)) Then
            .AddItem h
        End If
        
        MarsAutomaticDatum(1) = VerkaufAdoRec!Datum
    End With
    
    VerkaufAdoRec.MoveNext
Loop
'With lblMarsAnzahlRezepte
'    .Caption = CStr(lstMarsRezepte.ListCount) + " Rezepte"
'    .Visible = True
'End With
With lblMarsModus
    .Caption = "Rezept-KONTROLLE" + vbCrLf + CStr(lstMarsRezepte.ListCount) + " Rezepte"
    .Visible = True
End With

If (lstMarsRezepte.ListCount > 0) Then
    txtRezeptNr.text = lstMarsRezepte.List(0)
    txtRezeptNr.SetFocus
    cmdOk(0).Value = True
    MarsAutomaticModus = 1
    MarsRezeptZurückgestellt = False
End If

Call DefErrPop
End Sub

Sub MarsNextRezept(Optional Zurückgestellt As Boolean = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MarsNextRezept")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    
If (MarsRezeptZurückgestellt = False) Then
    h = Format(MarsAutomaticDatum(0), "DD.MM.YYYY HH:mm:ss")
    l& = WritePrivateProfileString("Rezeptkontrolle", "LetztDatum", h$, CurDir + "\Ocr2016.ini")
    If (Zurückgestellt) Then
        MarsRezeptZurückgestellt = True
    End If
End If

'Call TerminateProcess("OpRezept.exe")
l& = EntferneTask("OP-Rezept-Viewer")

SQLStr = "SELECT * FROM Verkauf"
SQLStr = SQLStr + " WHERE ((RezeptNr<>'') OR (PrivRezLaufNr>0))"
'SQLStr = SQLStr + " AND (Datum<='" + Left(h, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5, 2) + " 23:59" + "')"
SQLStr = SQLStr + " AND (Datum>'" + Format(MarsAutomaticDatum(1), "DD.MM.YYYY HH:mm:ss") + "')"
SQLStr = SQLStr + " ORDER BY Datum,LaufNr,ZeilenNr"
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
If (FabsErrf = 0) Then
    VerkaufAdoRec.MoveFirst
    Do
        If (VerkaufAdoRec.EOF) Then
            Exit Do
        End If
        
        h = Trim(CheckNullStr(VerkaufAdoRec!RezeptNr))
        If (h = "") Then
            h = CStr(CheckNullLong(VerkaufAdoRec!PrivRezlaufNr))
        End If
        
        With lstMarsRezepte
            For i = 0 To (.ListCount - 1)
                If (.List(i) = h) Then
                    .RemoveItem (i)
                    Exit For
                End If
            Next i
            
            If (TesteMarsRezept(h)) Then
                .AddItem h
            End If
        End With
        
        VerkaufAdoRec.MoveNext
    Loop
End If

'DoEvents
With lstMarsRezepte
    .RemoveItem (0)
    If (.ListCount > 0) Then
'        With lblMarsAnzahlRezepte
'            .Caption = CStr(lstMarsRezepte.ListCount) + " Rezepte"
'            .Visible = True
'        End With
        With lblMarsModus
            .Caption = "Rezept-KONTROLLE" + vbCrLf + CStr(lstMarsRezepte.ListCount) + " Rezepte"
        End With
        
        txtRezeptNr.text = .List(0)
        txtRezeptNr.SetFocus
        cmdOk(0).Value = True
    Else
'        With lblMarsAnzahlRezepte
'            .Caption = CStr(lstMarsRezepte.ListCount) + " Rezepte"
'            .Visible = False
'        End With
        With lblMarsModus
            .Caption = "Rezept-KONTROLLE" + vbCrLf
        End With
        MarsAutomaticModus = 0
        Call MessageBox("Keine weiteren zu kontrollierenden Rezepte", vbInformation)
    End If
End With

Call DefErrPop
End Sub

Function TesteMarsRezept(RezNr$) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TesteMarsRezept")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret As Boolean

ret = False
If (para.MARS) Then
    Dim i%, l%, ch$, TestlRezNr&, TestRezNr$, hStr$
    
    ret = True
    
    TestRezNr = RezNr
    l% = Len(TestRezNr)
    i% = 1
    Do
        If (i% > l%) Then Exit Do

        ch$ = Mid$(TestRezNr$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            i% = i% + 1
        Else
            TestRezNr$ = Left$(TestRezNr$, i% - 1) + Mid$(TestRezNr$, i% + 1)
            l% = Len(TestRezNr$)
        End If
    Loop
    If (l% > 13) Then TestRezNr$ = Left$(TestRezNr$, 13)

    TestlRezNr& = 1&
    If (l% < 13) Then
        If (l% <= 5) Then
            TestlRezNr& = Val(TestRezNr$)
        Else
            TestlRezNr& = 0
        End If
        If (TestlRezNr& <= 0) Then
            TesteMarsRezept = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(TestlRezNr&, "00000")
        Call PruefZiffer(hStr$)
        TestRezNr$ = hStr$
    End If
    
    
'    SQLStr = "SELECT * FROM Anmerkungen WHERE RezeptNr='" + TestRezNr + "'"
'    Set MarsAnmerkungenRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
'    If (MarsAnmerkungenRec.RecordCount > 0) Then
'        MarsRezeptAnmerkungen = Trim(CheckNullStr(MarsAnmerkungenRec!Anmerkungen))
'    End If
'
'    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr + "'"
'    Set RezepteRec = RezSpeicherDB.OpenRecordset(SQLStr$)
'    If (RezepteRec.RecordCount > 0) Then
'        h = CheckNullStr(RezepteRec!DruckDatum)
'        Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde BEREITS BEDRUCKT (" + Mid(h, 5, 2) + "." + Mid(h, 3, 2) + ".20" + Left(h, 2) + ") !", vbInformation)
'    End If
'
'    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr + "'"
'    Set MarsRezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
'    If (MarsRezepteRec.RecordCount > 0) Then
'        If (MarsModus = MARS_REZEPT_KONTROLLE) Then
'            Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde BEREITS FREIGEGEBEN!", vbInformation)
'        End If
'    Else
'        If (MarsModus = MARS_REZEPT_DRUCK) Then
'            Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde NOCH NICHT FREIGEGEBEN!", vbCritical)
'            RezeptHolenDB% = False
'            Call DefErrPop: Exit Function
'        End If
'    End If

    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr + "'"
    Set MarsRezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    If (MarsRezepteRec.RecordCount > 0) Then
        ret = False
    End If
End If

TesteMarsRezept = ret

Call DefErrPop
End Function

Private Sub TerminateProcess(app_exe As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TerminateProcess")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
Dim Process As Object
For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
    Process.Terminate
Next

Call DefErrPop
End Sub

Sub MarsSaveRezeptAnmerkungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MarsSaveRezeptAnmerkungen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
If (para.MARS) Then
    SQLStr = "DELETE * FROM Anmerkungen WHERE RezeptNr='" + RezNr + "'"
    MarsRezSpeicherDB.Execute (SQLStr$)
    MarsRezeptAnmerkungen = txtMarsRezeptAnmerkungen.text
    SQLStr = "INSERT INTO Anmerkungen (RezeptNr,Anmerkungen) VALUES ('" + RezNr + "','" + MarsRezeptAnmerkungen + "')"
    MarsRezSpeicherDB.Execute (SQLStr$)
    
'    Call TerminateProcess("OpRezept.exe")
    Dim l&
    l& = EntferneTask("OP-Rezept-Viewer")
End If

Call DefErrPop
End Sub
    

