VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmLiefStammdaten 
   AutoRedraw      =   -1  'True
   Caption         =   "Stammdaten von "
   ClientHeight    =   8775
   ClientLeft      =   -660
   ClientTop       =   225
   ClientWidth     =   12825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12825
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   12000
      Picture         =   "LiefStammdaten.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   12240
      Picture         =   "LiefStammdaten.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   12480
      Picture         =   "LiefStammdaten.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3480
      TabIndex        =   30
      Top             =   7320
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1920
      TabIndex        =   29
      Top             =   7320
      Width           =   1200
   End
   Begin TabDlg.SSTab tabStammdaten 
      Height          =   6765
      Left            =   1080
      TabIndex        =   31
      Top             =   360
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   11933
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 - Allgemein"
      TabPicture(0)   =   "LiefStammdaten.frx":0216
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeStammdaten(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - FIBU"
      TabPicture(1)   =   "LiefStammdaten.frx":0232
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeStammdaten(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Optionen"
      TabPicture(2)   =   "LiefStammdaten.frx":024E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeStammdaten(2)"
      Tab(2).Control(1)=   "tmrLiefStammdaten"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "&4 - Verbindungen"
      TabPicture(3)   =   "LiefStammdaten.frx":026A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeStammdaten(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5 - Rabatt-Tab."
      TabPicture(4)   =   "LiefStammdaten.frx":0286
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeStammdaten(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Ausnahmen"
      TabPicture(5)   =   "LiefStammdaten.frx":02A2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fmeStammdaten(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&7 - Sonstiges"
      TabPicture(6)   =   "LiefStammdaten.frx":02BE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fmeStammdaten(6)"
      Tab(6).ControlCount=   1
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   5655
         Index           =   6
         Left            =   -73680
         TabIndex        =   117
         Top             =   720
         Width           =   8175
         Begin VB.CheckBox chkSonstiges2 
            Caption         =   "&Fixabatte"
            Height          =   375
            Left            =   360
            TabIndex        =   142
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Frame fmeSonstiges2 
            Caption         =   "ALLE Angebote dieses Lieferanten"
            Height          =   1815
            Left            =   1560
            TabIndex        =   143
            Top             =   3600
            Width           =   5055
            Begin VB.TextBox txtSonstiges2 
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
               Left            =   3120
               TabIndex        =   145
               Text            =   "999.99"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtSonstiges2 
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
               Left            =   3120
               TabIndex        =   147
               Text            =   "999.99"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblSonstiges2 
               Caption         =   "Fixer &Rabatt in % (RX)"
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   144
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label lblSonstiges2 
               Caption         =   "Fixer &Rabatt in % (NonRX)"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   146
               Top             =   1080
               Width           =   2895
            End
         End
         Begin VB.ComboBox cboSonstiges 
            Height          =   315
            Index           =   0
            Left            =   2520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown-Liste
            TabIndex        =   135
            Top             =   0
            Width           =   1935
         End
         Begin VB.Frame fmeSonstiges 
            Caption         =   "GH-Angebote mit Menge=1"
            Height          =   2775
            Left            =   1680
            TabIndex        =   137
            Top             =   600
            Width           =   5175
            Begin VB.TextBox txtSonstiges 
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
               Left            =   3120
               TabIndex        =   139
               Text            =   "999.99"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   360
               Width           =   1455
            End
            Begin VB.ComboBox cboSonstiges 
               Height          =   315
               Index           =   1
               Left            =   3000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown-Liste
               TabIndex        =   141
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label lblSonstiges 
               Caption         =   "Fixer &Rabatt in %"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   138
               Top             =   360
               Width           =   2895
            End
            Begin VB.Label lblSonstiges 
               Caption         =   "&Bestellen bei Lieferant"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   140
               Top             =   1200
               Width           =   2415
            End
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Pool"
            Height          =   375
            Left            =   0
            TabIndex        =   136
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblSonstiges 
            Caption         =   "&Umsatz speichern bei Lieferant"
            Height          =   615
            Index           =   0
            Left            =   0
            TabIndex        =   134
            Top             =   0
            Width           =   4455
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   5655
         Index           =   5
         Left            =   -74400
         TabIndex        =   100
         Top             =   720
         Width           =   7815
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "&AM Rx"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   133
            Top             =   4440
            Width           =   2535
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4800
            MaxLength       =   5
            TabIndex        =   132
            Text            =   "Wwwww"
            Top             =   3960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3720
            MaxLength       =   5
            TabIndex        =   131
            Text            =   "Wwwww"
            Top             =   3960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   130
            Text            =   "Wwwww"
            Top             =   3960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   4680
            MaxLength       =   5
            TabIndex        =   129
            Text            =   "Wwwww"
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   3720
            MaxLength       =   5
            TabIndex        =   128
            Text            =   "Wwwww"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   9
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   110
            Text            =   "99"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   8
            Left            =   5880
            MaxLength       =   2
            TabIndex        =   121
            Text            =   "99"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   7
            Left            =   5040
            MaxLength       =   2
            TabIndex        =   120
            Text            =   "99"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   6
            Left            =   4200
            MaxLength       =   2
            TabIndex        =   119
            Text            =   "99"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   5
            Left            =   3360
            MaxLength       =   2
            TabIndex        =   118
            Text            =   "99"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   107
            Text            =   "99"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   5880
            MaxLength       =   2
            TabIndex        =   116
            Text            =   "99"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   5040
            MaxLength       =   2
            TabIndex        =   115
            Text            =   "99"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   4200
            MaxLength       =   2
            TabIndex        =   114
            Text            =   "99"
            Top             =   960
            Width           =   735
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "&Kühlartikel"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   108
            Top             =   0
            Width           =   2535
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "&Sonderangebote"
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   109
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "&Betäubungsmittel"
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   111
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   " &Warengruppen"
            Height          =   375
            Index           =   3
            Left            =   0
            TabIndex        =   112
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "bis Bestell&menge"
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   122
            Top             =   2400
            Width           =   2535
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "bis &Zeilenwert"
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   124
            Top             =   2880
            Width           =   2535
         End
         Begin VB.CheckBox chkAusnahmen 
            Caption         =   "&Hersteller"
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   126
            Top             =   3480
            Width           =   2535
         End
         Begin VB.TextBox txtAusnahmenBM 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            MaxLength       =   3
            TabIndex        =   123
            Text            =   "9999"
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenZW 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   125
            Text            =   "9999.99"
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenWg 
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   3360
            MaxLength       =   2
            TabIndex        =   113
            Text            =   "99"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtAusnahmenHerst 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2640
            MaxLength       =   5
            TabIndex        =   127
            Text            =   "Wwwww"
            Top             =   3480
            Width           =   735
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   4815
         Index           =   4
         Left            =   -74520
         TabIndex        =   94
         Top             =   840
         Width           =   5175
         Begin VB.CommandButton cmdF2 
            Caption         =   "Einfügen (F2)"
            Height          =   450
            Left            =   1920
            TabIndex        =   105
            Top             =   1320
            Width           =   1200
         End
         Begin VB.CommandButton cmdF5 
            Caption         =   "Entfernen (F5)"
            Height          =   450
            Left            =   1920
            TabIndex        =   106
            Top             =   1920
            Width           =   1200
         End
         Begin VB.OptionButton optRabattTabelle 
            Caption         =   "&Einkaufswert-Tabelle"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   101
            Top             =   0
            Width           =   2895
         End
         Begin VB.OptionButton optRabattTabelle 
            Caption         =   "&BM-/Zeilenrabatt - Tabelle"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   102
            Top             =   360
            Width           =   2895
         End
         Begin MSFlexGridLib.MSFlexGrid flxRabattTabelle 
            Height          =   2040
            Index           =   0
            Left            =   0
            TabIndex        =   103
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   3598
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483633
            BackColorBkg    =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin MSFlexGridLib.MSFlexGrid flxRabattTabelle 
            Height          =   2040
            Index           =   1
            Left            =   840
            TabIndex        =   104
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   3598
            _Version        =   393216
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483633
            BackColorBkg    =   -2147483633
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   5055
         Index           =   3
         Left            =   -74520
         TabIndex        =   84
         Top             =   960
         Width           =   9615
         Begin VB.CheckBox chkStammdaten4 
            Caption         =   "M&odem"
            Height          =   495
            Index           =   0
            Left            =   1200
            TabIndex        =   89
            Top             =   0
            Width           =   855
         End
         Begin VB.TextBox txtStammdaten4 
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
            Left            =   2640
            TabIndex        =   90
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   0
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten4 
            Caption         =   "e&Mail"
            Height          =   495
            Index           =   1
            Left            =   1080
            TabIndex        =   91
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtStammdaten4 
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
            Left            =   2520
            TabIndex        =   92
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten4 
            Caption         =   "&Fax"
            Height          =   495
            Index           =   2
            Left            =   960
            TabIndex        =   93
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtStammdaten4 
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
            Index           =   2
            Left            =   2400
            TabIndex        =   95
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1200
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten4 
            Caption         =   "automat. Aus&druck"
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   97
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtStammdaten4 
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
            Index           =   3
            Left            =   3960
            TabIndex        =   99
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3120
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten4 
            Caption         =   "&Computer-Fax"
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   96
            Top             =   1140
            Width           =   3015
         End
         Begin VB.Label lblStammdaten4 
            Caption         =   "&Angebote aus dem Internet unter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   98
            Top             =   3240
            Width           =   4095
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   5175
         Index           =   1
         Left            =   -74760
         TabIndex        =   82
         Top             =   840
         Width           =   9735
         Begin VB.CheckBox chkStammdaten2 
            Caption         =   "&Depot-Pflicht"
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   53
            Top             =   4680
            Width           =   3015
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   8
            Left            =   4440
            TabIndex        =   52
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   4680
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Left            =   4560
            TabIndex        =   34
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            TabIndex        =   36
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   540
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   2
            Left            =   4560
            TabIndex        =   38
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1140
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   3
            Left            =   4440
            TabIndex        =   41
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   2220
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   5
            Left            =   4440
            TabIndex        =   46
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3300
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   4
            Left            =   4440
            TabIndex        =   43
            TabStop         =   0   'False
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   2640
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   6
            Left            =   4440
            TabIndex        =   48
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3660
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten2 
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
            Index           =   7
            Left            =   4440
            TabIndex        =   50
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   4140
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten2 
            Caption         =   "Bank&einzug"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Width           =   5535
         End
         Begin VB.ComboBox cboStammdaten2 
            Height          =   315
            Left            =   5280
            Style           =   2  'Dropdown-Liste
            TabIndex        =   44
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "Lieferfrist in &Tagen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   51
            Top             =   4560
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "Bank-&Kontonummer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "&Bank/BLZ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   35
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "Konto &Vorschlag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "&FIBU-Kontonummer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   40
            Top             =   2160
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "&Wareneingang netto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   45
            Top             =   3240
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "&Drogerie/Vertreter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   42
            Top             =   2760
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "&Skonto Vorschlag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   47
            Top             =   3600
            Width           =   4095
         End
         Begin VB.Label lblStammdaten2 
            Caption         =   "Saldo offene &Posten"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   49
            Top             =   4080
            Width           =   4095
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   5655
         Index           =   0
         Left            =   360
         TabIndex        =   78
         Top             =   960
         Width           =   9855
         Begin VB.OptionButton optStammdaten 
            Caption         =   "Automatik&bestellung"
            Height          =   375
            Index           =   2
            Left            =   6120
            TabIndex        =   18
            Top             =   2160
            Width           =   3375
         End
         Begin VB.OptionButton optStammdaten 
            Caption         =   "Zuordnungen &gelten"
            Height          =   375
            Index           =   1
            Left            =   6240
            TabIndex        =   17
            Top             =   1560
            Width           =   3375
         End
         Begin VB.OptionButton optStammdaten 
            Caption         =   "Zuordnungen gelten ni&cht"
            Height          =   375
            Index           =   0
            Left            =   6240
            TabIndex        =   16
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   11
            Left            =   5040
            TabIndex        =   28
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   5280
            Width           =   375
         End
         Begin VB.CheckBox chkStammdaten 
            Caption         =   "Direkt&bestellung automatisch"
            Height          =   375
            Index           =   1
            Left            =   5880
            TabIndex        =   15
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtStammdaten 
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
            Left            =   4440
            TabIndex        =   1
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   60
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            TabIndex        =   3
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   420
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   2
            Left            =   4440
            TabIndex        =   5
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   900
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   3
            Left            =   4440
            TabIndex        =   7
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   4
            Left            =   4440
            TabIndex        =   9
            Text            =   "WWWWWW"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1620
            Width           =   1455
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   5
            Left            =   4440
            TabIndex        =   11
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   2100
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   6
            Left            =   4440
            TabIndex        =   13
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   2460
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   7
            Left            =   4800
            TabIndex        =   20
            Text            =   "99999999999999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3660
            Width           =   2775
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   8
            Left            =   4800
            TabIndex        =   22
            Text            =   "9999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   4020
            Width           =   855
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   9
            Left            =   4920
            TabIndex        =   24
            Text            =   "9999999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   4380
            Width           =   1695
         End
         Begin VB.TextBox txtStammdaten 
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
            Index           =   10
            Left            =   4800
            TabIndex        =   26
            Text            =   "9999999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   4740
            Width           =   1935
         End
         Begin VB.CheckBox chkStammdaten 
            Caption         =   "&Direktlieferant"
            Height          =   375
            Index           =   0
            Left            =   6000
            TabIndex        =   14
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Zusatzaufwand (0-100% des AEP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   27
            Top             =   5280
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Name/Adresse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Kurzbezeichnung"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   8
            Top             =   1620
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Warenbezeichnung"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   10
            Top             =   2100
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Vertretername"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   12
            Top             =   2460
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "&Telefon"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   19
            Top             =   3660
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "D&urchwahl"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   21
            Top             =   4020
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "IDF-Nummer &Apotheke"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   23
            Top             =   4380
            Width           =   4095
         End
         Begin VB.Label lblStammdaten 
            Caption         =   "IDF-Nummer &Lieferant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   25
            Top             =   4740
            Width           =   4095
         End
      End
      Begin VB.Frame fmeStammdaten 
         Caption         =   "Frame1"
         Height          =   6375
         Index           =   2
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   9375
         Begin VB.ComboBox cboStammdaten3 
            Height          =   315
            Index           =   1
            Left            =   4080
            Style           =   2  'Dropdown-Liste
            TabIndex        =   81
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdHersteller 
            Caption         =   "H&inzufügen"
            Height          =   375
            Left            =   6240
            TabIndex        =   67
            Top             =   960
            Width           =   2655
         End
         Begin VB.ListBox lstStammdaten3 
            Height          =   450
            Left            =   6360
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   15
            Left            =   4320
            TabIndex        =   59
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   5880
            Width           =   375
         End
         Begin VB.CommandButton cmdZuordnungen 
            Caption         =   "&Artikelzuordnungen ..."
            Height          =   375
            Left            =   5160
            TabIndex        =   70
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox txtStammdaten3 
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
            Left            =   4440
            TabIndex        =   55
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   60
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            TabIndex        =   57
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   540
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   3
            Left            =   4440
            TabIndex        =   63
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1140
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   4
            Left            =   4440
            TabIndex        =   65
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1620
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   6
            Left            =   4560
            TabIndex        =   72
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   2700
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   7
            Left            =   4560
            TabIndex        =   74
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3180
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   8
            Left            =   4560
            TabIndex        =   76
            Text            =   "9999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3900
            Width           =   375
         End
         Begin VB.ComboBox cboStammdaten3 
            Height          =   315
            Index           =   0
            Left            =   4080
            Style           =   2  'Dropdown-Liste
            TabIndex        =   79
            Top             =   4560
            Width           =   1935
         End
         Begin VB.Frame fmeStammdaten3 
            Caption         =   "Tem&porär"
            Height          =   2775
            Left            =   6360
            TabIndex        =   83
            Top             =   2760
            Width           =   2775
            Begin VB.TextBox txtStammdaten3 
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
               Index           =   9
               Left            =   960
               TabIndex        =   85
               Text            =   "999"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtStammdaten3 
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
               Index           =   10
               Left            =   960
               TabIndex        =   86
               Text            =   "999"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox txtStammdaten3 
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
               Index           =   11
               Left            =   1080
               TabIndex        =   87
               Text            =   "999"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   1440
               Width           =   375
            End
            Begin VB.TextBox txtStammdaten3 
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
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
               Index           =   12
               Left            =   960
               TabIndex        =   88
               Text            =   "999"
               ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
               Top             =   2040
               Width           =   375
            End
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   2
            Left            =   5520
            TabIndex        =   61
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   900
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   13
            Left            =   5040
            TabIndex        =   68
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtStammdaten3 
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
            Index           =   14
            Left            =   5640
            TabIndex        =   69
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "&Preisbasis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   80
            Top             =   5160
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3Zus 
            Caption         =   "Prozent erhöhen, wenn unerreicht"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   148
            Top             =   5880
            Width           =   1935
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Bevorratungszeitraum um"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1320
            TabIndex        =   58
            Top             =   5880
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Direktbezug möglich ab &BM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "&Mindestwert für Direktbezug in EUR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   56
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Verr.&Kosten pro Lieferung in EUR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   62
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Lieferant für &Hersteller"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   64
            Top             =   1560
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Bevorratungs&zeitraum in Tagen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   71
            Top             =   2640
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "&Valutastellung in Tagen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   73
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "&Fakturenrabatt in %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   75
            Top             =   3840
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "Fakturen&rabatt als"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   77
            Top             =   4500
            Width           =   4095
         End
         Begin VB.Label lblStammdaten3 
            Caption         =   "&Mindestbewertung für Direktbezug"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   60
            Top             =   840
            Width           =   4095
         End
      End
      Begin VB.Timer tmrLiefStammdaten 
         Interval        =   100
         Left            =   -65160
         Top             =   1080
      End
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   1800
      TabIndex        =   152
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   3600
      TabIndex        =   153
      Top             =   7920
      Width           =   1095
      _ExtentX        =   4551
      _ExtentY        =   1058
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmLiefStammdaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TxtInd%
Dim VerglStr$(4)
Dim iEditModus%

Dim OrgBevorratungsZeit%
Dim OrgFakturenRabatt#

Private Const DefErrModul = "LIEFSTAMMDATEN.FRM"

Private Sub chkStammdaten_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkStammdaten_Click")
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

If (index = 1) Then
    With chkStammdaten(index)
        If (.Value) Then
            .ForeColor = chkStammdaten(0).ForeColor
        Else
            .ForeColor = vbRed
        End If
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub chkStammdaten4_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkStammdaten4_Click")
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
If (index < 3) Then txtStammdaten4(index).Enabled = chkStammdaten4(index).Value
If (index = 2) Then chkStammdaten4(4).Enabled = chkStammdaten4(index).Value
Call clsError.DefErrPop
End Sub

Private Sub chkSonstiges_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkSonstiges_Click")
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

fmeSonstiges.Visible = chkSonstiges.Value

Call clsError.DefErrPop
End Sub

Private Sub chkSonstiges2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkSonstiges2_Click")
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

fmeSonstiges2.Visible = chkSonstiges2.Value

Call clsError.DefErrPop
End Sub

Private Sub cmdEsc_Click()
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

EditErg% = False
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdHersteller_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdHersteller_Click")
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
Dim i%, neu%
Dim h$

h$ = UCase(Trim(txtStammdaten3(4).text))
If (h$ <> "") Then
    With lstStammdaten3
        neu% = True
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            If (.text = h$) Then
                neu% = False
                Exit For
            End If
        Next i%
        If (neu%) Then
            .AddItem h$
        End If
        txtStammdaten3(4).text = ""
        txtStammdaten3(4).SetFocus
    End With
Else
    With lstStammdaten3
        If (.ListIndex >= 0) Then
            .RemoveItem .ListIndex
        End If
        .SetFocus
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdOk_Click()
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
Dim ind%

If (ActiveControl.Name = cmdOk.Name) Or (ActiveControl.Name = nlcmdOk.Name) Then
    Call StammdatenSpeichern
    Call SpeicherRabattTabelle
    EditErg% = True
    Unload Me
ElseIf (ActiveControl.Name = flxRabattTabelle(0).Name) Then
    ind% = ActiveControl.index
    If (ind% = 1) And (flxRabattTabelle(1).col = 2) Then
        Call EditSchwellwerteLst
    ElseIf (ind% = 1) And (flxRabattTabelle(1).col = 5) Then
        Call EditSchwellwerteLstMulti
    Else
        Call EditSchwellwerteTxt
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdZuordnungen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdZuordnungen_Click")
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

frmArtikelZuordnung.Show 1

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
Dim i%, j%, l%, k%, lInd%, MaxWi%, spBreite%, ind%, iLief%(1)
Dim Breite%, Breite2%, Hoehe%, Hoehe2%, xpos%, ydiff%, val2%
Dim iAdd%, iAdd2%, x%, y%, wi%
Dim val1&
Dim dVal#
Dim h$, h2$(1), FormStr$
Dim c As Control

VerglStr$(0) = "<"
VerglStr$(1) = "<="
VerglStr$(2) = "="
VerglStr$(3) = ">="
VerglStr$(4) = ">"

For i% = 0 To 11
    h$ = " " + lblStammdaten(i%).Caption
    lblStammdaten(i%).Caption = h$
Next i%
For i% = 0 To 8
    h$ = " " + lblStammdaten2(i%).Caption
    lblStammdaten2(i%).Caption = h$
Next i%
On Error Resume Next
For i% = 0 To 11
    h$ = " " + lblStammdaten3(i%).Caption
    lblStammdaten3(i%).Caption = h$
Next i%
On Error GoTo DefErr
For i% = 0 To 0
    h$ = " " + lblStammdaten4(i%).Caption
    lblStammdaten4(i%).Caption = h$
Next i%

With cboStammdaten2
    .AddItem "Drogerie"
    .AddItem "Vertreter"
End With
With cboStammdaten3(0)
    .AddItem "Bar - alle Artikel"
    .AddItem "Bar - ohne NR"
    .AddItem "Natural-Rabatt"
End With
With cboStammdaten3(1)
    .AddItem "Taxe-EK"
    .AddItem "HAP"
    .AddItem "HAP+70ct"
End With

Call wPara1.InitFont(Me)

For i% = 0 To 3
    txtStammdaten(i%).text = String(30, "X")
Next i%
txtStammdaten(4).text = String(6, "W")
txtStammdaten(5).text = String(16, "X")
txtStammdaten(6).text = String(15, "X")
txtStammdaten(7).text = String(14, "9")
txtStammdaten(8).text = String(4, "9")
txtStammdaten(9).text = String(7, "9")
txtStammdaten(10).text = String(7, "9")
txtStammdaten(11).text = String(3, "9")

For i% = 0 To 11
    txtStammdaten(i%).MaxLength = Len(txtStammdaten(i%).text)
Next i%

txtStammdaten2(0).text = String(14, "9")
txtStammdaten2(1).text = String(8, "9")
txtStammdaten2(2).text = String(6, "9")
txtStammdaten2(3).text = String(6, "9")
For i% = 5 To 7
    txtStammdaten2(i%).text = "9999999.99"
Next i%
txtStammdaten2(8).text = String(3, "9")

For i% = 0 To 8
    txtStammdaten2(i%).MaxLength = Len(txtStammdaten2(i%).text)
Next i%

txtStammdaten3(0).text = String(3, "9")
txtStammdaten3(1).text = String(8, "9")
txtStammdaten3(2).text = String(8, "9")
txtStammdaten3(3).text = String(6, "9")
txtStammdaten3(4).text = String(5, "W")
'txtStammdaten3(5).text = String(3, "9")
txtStammdaten3(6).text = String(3, "9")
txtStammdaten3(7).text = String(3, "9")
txtStammdaten3(8).text = String(4, "9")
txtStammdaten3(9).text = String(3, "9")
txtStammdaten3(10).text = String(3, "9")
txtStammdaten3(11).text = String(3, "9")
txtStammdaten3(12).text = String(6, "9")
txtStammdaten3(13).text = String(5, "W")
txtStammdaten3(14).text = String(5, "W")
txtStammdaten3(15).text = String(3, "9")

On Error Resume Next
For i% = 0 To 15
    txtStammdaten3(i%).MaxLength = Len(txtStammdaten3(i%).text)
Next i%
On Error GoTo DefErr

txtStammdaten4(0).text = String(15, "9")
txtStammdaten4(1).text = String(20, "W")
txtStammdaten4(2).text = String(15, "9")
txtStammdaten4(3).text = String(20, "W")

For i% = 0 To 3
    txtStammdaten4(i%).MaxLength = Len(txtStammdaten4(i%).text)
Next i%

For i% = 1 To 9
'    Load txtAusnahmenWg(i%)
    txtAusnahmenWg(i%).text = txtAusnahmenWg(0).text
    txtAusnahmenWg(i%).Visible = True
    txtAusnahmenWg(i%).TabIndex = txtAusnahmenWg(i% - 1).TabIndex + 1
Next i%
For i% = 1 To 5
'    Load txtAusnahmenHerst(i%)
    txtAusnahmenHerst(i%).text = txtAusnahmenHerst(0).text
    txtAusnahmenHerst(i%).Visible = True
    txtAusnahmenHerst(i%).TabIndex = txtAusnahmenHerst(i% - 1).TabIndex + 1
Next i%

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
'        c.text = ""
    End If
Next
On Error GoTo DefErr

Call StammdatenBefuellen
Me.Caption = Me.Caption + RTrim(Lif1.kurz)


'tabStammdaten.Left = wPara1.LinksX
'tabStammdaten.Top = wPara1.TitelY

Font.Bold = False   ' True

tabStammdaten.Tab = 0

MaxWi% = 0
For i% = 0 To 11
    lblStammdaten(i%).Left = wPara1.LinksX
    wi% = lblStammdaten(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
txtStammdaten(0).Left = lblStammdaten(0).Left + MaxWi% + 300
For i% = 1 To 11
    txtStammdaten(i%).Left = txtStammdaten(i% - 1).Left
Next i%



lblStammdaten(0).Top = 3 * wPara1.TitelY

ydiff% = (txtStammdaten(0).Height - lblStammdaten(0).Height) / Screen.TwipsPerPixelY
ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
txtStammdaten(0).Top = lblStammdaten(0).Top - ydiff%

For i% = 1 To 3
    txtStammdaten(i%).Top = txtStammdaten(i% - 1).Top + txtStammdaten(i% - 1).Height
    lblStammdaten(i%).Top = txtStammdaten(i%).Top
Next i%

lblStammdaten(4).Top = txtStammdaten(3).Top + txtStammdaten(3).Height + 105
lblStammdaten(5).Top = lblStammdaten(4).Top + lblStammdaten(4).Height + 210
lblStammdaten(6).Top = lblStammdaten(5).Top + lblStammdaten(5).Height + 105

chkStammdaten(0).Left = txtStammdaten(6).Left + txtStammdaten(6).Width + 600
chkStammdaten(0).Top = lblStammdaten(6).Top
'chkStammdaten(1).Left = chkStammdaten(0).Left
'chkStammdaten(1).Top = chkStammdaten(0).Top + chkStammdaten(0).Height + 90
chkStammdaten(1).Visible = False
For i% = 0 To 2
    optStammdaten(i%).Left = chkStammdaten(0).Left
Next i%
optStammdaten(0).Top = chkStammdaten(0).Top + chkStammdaten(0).Height + 180
For i% = 1 To 2
    optStammdaten(i%).Top = optStammdaten(i% - 1).Top + optStammdaten(i% - 1).Height
Next i%

lblStammdaten(7).Top = lblStammdaten(6).Top + lblStammdaten(6).Height + 210
lblStammdaten(8).Top = lblStammdaten(7).Top + lblStammdaten(7).Height + 105

lblStammdaten(9).Top = lblStammdaten(8).Top + lblStammdaten(8).Height + 210
lblStammdaten(10).Top = lblStammdaten(9).Top + lblStammdaten(9).Height + 105

lblStammdaten(11).Top = lblStammdaten(10).Top + lblStammdaten(10).Height + 210

For i% = 4 To 11
    txtStammdaten(i%).Top = lblStammdaten(i%).Top - ydiff%
Next i%


Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

Hoehe% = txtStammdaten(11).Top + txtStammdaten(11).Height + 180
'Breite% = txtStammdaten(0).Left + txtStammdaten(0).Width + wPara1.LinksX
Breite% = optStammdaten(0).Left + optStammdaten(0).Width + wPara1.LinksX

'-------------------------


tabStammdaten.Tab = 1

MaxWi% = 0
For i% = 0 To 8
    lblStammdaten2(i%).Left = wPara1.LinksX
    wi% = lblStammdaten2(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
txtStammdaten2(0).Left = lblStammdaten2(0).Left + MaxWi% + 300
For i% = 1 To 8
    txtStammdaten2(i%).Left = txtStammdaten2(i% - 1).Left
Next i%
txtStammdaten2(4).Visible = False



lblStammdaten2(0).Top = 3 * wPara1.TitelY

ydiff% = (txtStammdaten2(0).Height - lblStammdaten2(0).Height) / 2
txtStammdaten2(0).Top = lblStammdaten2(0).Top - ydiff%

For i% = 1 To 2
    lblStammdaten2(i%).Top = lblStammdaten2(i% - 1).Top + lblStammdaten2(i% - 1).Height + 105
Next i%
chkStammdaten2(0).Left = lblStammdaten2(0).Left
chkStammdaten2(0).Top = lblStammdaten2(2).Top + lblStammdaten2(2).Height + 105
lblStammdaten2(3).Top = chkStammdaten2(0).Top + chkStammdaten2(0).Height + 210
lblStammdaten2(4).Top = lblStammdaten2(3).Top + lblStammdaten2(3).Height + 240
lblStammdaten2(5).Top = lblStammdaten2(4).Top + lblStammdaten2(4).Height + 240
For i% = 6 To 8
    lblStammdaten2(i%).Top = lblStammdaten2(i% - 1).Top + lblStammdaten2(i% - 1).Height + 105
Next i%

For i% = 0 To 8
    txtStammdaten2(i%).Top = lblStammdaten2(i%).Top - ydiff%
Next i%

With cboStammdaten2
    .Left = txtStammdaten2(4).Left
    .Top = txtStammdaten2(4).Top
End With

chkStammdaten2(1).Left = lblStammdaten2(0).Left
chkStammdaten2(1).Top = lblStammdaten2(8).Top + lblStammdaten2(8).Height + 105

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

'Hoehe2% = txtStammdaten2(7).Top + txtStammdaten2(7).Height + 180
Hoehe2% = chkStammdaten2(1).Top + chkStammdaten2(1).Height + 180
Breite2% = txtStammdaten2(0).Left + txtStammdaten2(0).Width + wPara1.LinksX
If (Hoehe2% > Hoehe%) Then Hoehe% = Hoehe2%
If (Breite2% > Breite%) Then Breite% = Breite2%

'--------------------------

tabStammdaten.Tab = 2

MaxWi% = 0
For i% = 0 To 4
    lblStammdaten3(i%).Left = wPara1.LinksX
    wi% = lblStammdaten3(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
txtStammdaten3(0).Left = lblStammdaten3(0).Left + MaxWi% + 300
For i% = 1 To 4
    txtStammdaten3(i%).Left = txtStammdaten3(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 6 To 9
    lblStammdaten3(i%).Left = wPara1.LinksX
    wi% = lblStammdaten3(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
lblStammdaten3(11).Left = wPara1.LinksX
txtStammdaten3(6).Left = lblStammdaten3(0).Left + MaxWi% + 300
For i% = 7 To 8
    txtStammdaten3(i%).Left = txtStammdaten3(i% - 1).Left
Next i%


lblStammdaten3(0).Top = 3 * wPara1.TitelY

ydiff% = (txtStammdaten3(0).Height - lblStammdaten3(0).Height) / 2
txtStammdaten3(0).Top = lblStammdaten3(0).Top - ydiff%

For i% = 1 To 4
    lblStammdaten3(i%).Top = lblStammdaten3(i% - 1).Top + lblStammdaten3(i% - 1).Height + 150
    If (i% = 2) Then
        lblStammdaten3(i%).Top = lblStammdaten3(i%).Top + lblStammdaten3(i%).Height + 150
    End If
Next i%
lblStammdaten3(10).Top = lblStammdaten3(1).Top + lblStammdaten3(1).Height + 90

'lblStammdaten3(5).Top = lblStammdaten3(5).Top + 450

'chkStammdaten3.Left = lblStammdaten3(0).Left
'chkStammdaten3.Top = lblStammdaten3(5).Top + lblStammdaten3(5).Height + 210

'lblStammdaten3(6).Top = chkStammdaten3.Top + chkStammdaten3.Height + 300
lblStammdaten3(6).Top = lblStammdaten3(4).Top + lblStammdaten3(4).Height + 150 + 720
For i% = 7 To 9
    lblStammdaten3(i%).Top = lblStammdaten3(i% - 1).Top + lblStammdaten3(i% - 1).Height + 105
Next i%

On Error Resume Next
For i% = 0 To 8
    txtStammdaten3(i%).Top = lblStammdaten3(i%).Top - ydiff%
Next i%
On Error GoTo DefErr

txtStammdaten3(15).Top = lblStammdaten3(10).Top - ydiff%
txtStammdaten3(15).Left = txtStammdaten3(1).Left
lblStammdaten3(10).Left = txtStammdaten3(15).Left - lblStammdaten3(10).Width - 90
lblStammdaten3Zus.Top = lblStammdaten3(10).Top
lblStammdaten3Zus.Left = txtStammdaten3(15).Left + txtStammdaten3(15).Width + 90

With cboStammdaten3(0)
    .Left = txtStammdaten3(8).Left
    .Top = lblStammdaten3(9).Top
End With

For i% = 9 To 12
    txtStammdaten3(i%).Left = 2 * wPara1.LinksX
Next i%
ydiff% = txtStammdaten3(7).Top - txtStammdaten3(6).Top
txtStammdaten3(9).Top = 2 * wPara1.TitelY - 15 '+ 60
For i% = 10 To 12
    txtStammdaten3(i%).Top = txtStammdaten3(i% - 1).Top + ydiff%
'    txtStammdaten3(i%).Top = txtStammdaten3(i% - 1).Top + txtStammdaten3(i% - 1).Height + 150
Next i%

With lblStammdaten3(11)
    .Top = lblStammdaten3(9).Top + lblStammdaten3(9).Height + 105
End With
With cboStammdaten3(1)
    .Left = txtStammdaten3(8).Left
    .Top = lblStammdaten3(11).Top
End With

For i% = 13 To 14
    txtStammdaten3(i%).Top = txtStammdaten3(4).Top
    txtStammdaten3(i%).Visible = False
Next i%
txtStammdaten3(13).Left = txtStammdaten3(4).Left + txtStammdaten3(4).Width + 90
txtStammdaten3(14).Left = txtStammdaten3(13).Left + txtStammdaten3(13).Width + 90

With lstStammdaten3
    .Top = txtStammdaten3(4).Top
    .Left = txtStammdaten3(4).Left + txtStammdaten3(4).Width + 90
    .Width = txtStammdaten3(4).Width + TextWidth("wwww")
End With

With fmeStammdaten3
    .Left = cboStammdaten3(0).Left + cboStammdaten3(0).Width + 300
    .Top = txtStammdaten3(6).Top - txtStammdaten3(6).Height + 300  '+ 45 '- 45
    .Width = txtStammdaten3(12).Left + txtStammdaten3(12).Width + 2 * wPara1.LinksX
    .Height = txtStammdaten3(12).Top + txtStammdaten3(12).Height + wPara1.TitelY
End With

With cmdHersteller
    .Width = TextWidth(cmdZuordnungen.Caption) + 150
    .Height = wPara1.ButtonY
    .Top = lstStammdaten3.Top - 75
    .Left = lstStammdaten3.Left + lstStammdaten3.Width + 150
End With

With cmdZuordnungen
    .Width = cmdHersteller.Width
    .Height = wPara1.ButtonY
    .Top = cmdHersteller.Top + cmdHersteller.Height + 45
    .Left = cmdHersteller.Left
    lstStammdaten3.Height = .Top + .Height - lstStammdaten3.Top
End With


Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

'Hoehe2% = lblStammdaten3(9).Top + lblStammdaten3(9).Height + 180
Hoehe2% = fmeStammdaten3.Top + fmeStammdaten3.Height + 180
Breite2% = fmeStammdaten3.Left + fmeStammdaten3.Width + wPara1.LinksX
If (Hoehe2% > Hoehe%) Then Hoehe% = Hoehe2%
If (Breite2% > Breite%) Then Breite% = Breite2%


'--------------------------

tabStammdaten.Tab = 3

MaxWi% = 0
For i% = 0 To 2
    chkStammdaten4(i%).Left = wPara1.LinksX
    wi% = chkStammdaten4(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
txtStammdaten4(0).Left = lblStammdaten4(0).Left + MaxWi% + 300
For i% = 1 To 2
    txtStammdaten4(i%).Left = txtStammdaten4(i% - 1).Left
Next i%



chkStammdaten4(0).Top = 3 * wPara1.TitelY

ydiff% = (txtStammdaten4(0).Height - chkStammdaten4(0).Height) / 2
txtStammdaten4(0).Top = chkStammdaten4(0).Top - ydiff%

For i% = 1 To 2
    chkStammdaten4(i%).Top = chkStammdaten4(i% - 1).Top + chkStammdaten4(i% - 1).Height + 150
Next i%

chkStammdaten4(3).Left = chkStammdaten4(0).Left
chkStammdaten4(3).Top = chkStammdaten4(2).Top + chkStammdaten4(2).Height + 450

lblStammdaten4(0).Left = chkStammdaten4(0).Left
lblStammdaten4(0).Top = chkStammdaten4(3).Top + chkStammdaten4(3).Height + 450

For i% = 0 To 2
    txtStammdaten4(i%).Top = chkStammdaten4(i%).Top - ydiff%
Next i%
txtStammdaten4(3).Left = lblStammdaten4(0).Left + lblStammdaten4(0).Width + 300
txtStammdaten4(3).Top = lblStammdaten4(0).Top - ydiff%

chkStammdaten4(4).Top = chkStammdaten4(2).Top
chkStammdaten4(4).Left = txtStammdaten4(2).Left + txtStammdaten4(2).Width + 450

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

Hoehe2% = txtStammdaten4(3).Top + txtStammdaten4(3).Height + 180
Breite2% = txtStammdaten4(3).Left + txtStammdaten4(3).Width + wPara1.LinksX
If (Hoehe2% > Hoehe%) Then Hoehe% = Hoehe2%
If (Breite2% > Breite%) Then Breite% = Breite2%


'--------------------------

tabStammdaten.Tab = 4

Font.Name = wPara1.FontName(0)
Font.Size = wPara1.FontSize(0)

For i% = 0 To 1
    optRabattTabelle(i%).Left = wPara1.LinksX
Next i%
optRabattTabelle(0).Top = 3 * wPara1.TitelY   '2 * wPara1.TitelY
For i% = 1 To 1
    optRabattTabelle(i%).Top = optRabattTabelle(i% - 1).Top + optRabattTabelle(i% - 1).Height + 60
Next i%

With flxRabattTabelle(1)
    .Cols = 6
    .Rows = 6
    .FixedRows = 1
    .FixedCols = 0
    .row = 1
    
    .Top = optRabattTabelle(1).Top + optRabattTabelle(1).Height + 210
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * .Rows + 90
        
    .FormatString = ">AbBM|>BisBM|>Vgl|>zWert|>Rabatt (%)|<Gruppen"
    .SelectionMode = flexSelectionFree
    
    .ColWidth(0) = TextWidth("AbBM  ")
    .ColWidth(1) = TextWidth("BisBM  ")
    .ColWidth(2) = TextWidth("Vgl  ")
    .ColWidth(3) = TextWidth("999999999  ")
    .ColWidth(4) = TextWidth("Rabatt (%)    ")
    .ColWidth(5) = TextWidth(String(20, "X"))

    spBreite% = 0
    For i% = 0 To (.Cols - 1)
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    For i% = 0 To 4
        val2% = LifZus1.AbBM(i%)
        If (val2% > 0) Then .TextMatrix(i% + 1, 0) = Format(val2%, "0")
        val2% = LifZus1.BisBM(i%)
        If (val2% > 0) Then .TextMatrix(i% + 1, 1) = Format(val2%, "0")
        val2% = LifZus1.Vergleich(i%)
'        If (val2% > 0) Then .TextMatrix(i% + 1, 2) = Format(val2%, "0")
        If (val2% < 5) Then .TextMatrix(i% + 1, 2) = VerglStr$(val2%)
        val1& = LifZus1.zWert(i%)
        If (val1& > 0) Then .TextMatrix(i% + 1, 3) = Format(val1&, "0")
        dVal# = LifZus1.BmEkRabatt(i%)
        If (dVal# > 0) Then .TextMatrix(i% + 1, 4) = Format(dVal#, "0.00")
        
        val2% = LifZus1.BmEkTabGruppe%(i%)
        k% = 1
        h$ = ""
        For j% = 1 To 10
            If (val2% And k%) Then
                h$ = h$ + Format(j% - 1, "0")
            End If
            k% = k% * 2
        Next j%
        .TextMatrix(i% + 1, 5) = h$
    Next i%
End With

With flxRabattTabelle(0)
    .Cols = 2
    .Rows = 6
    .FixedRows = 1
    .FixedCols = 0
    .row = 1
    
    .Top = flxRabattTabelle(1).Top
    .Left = flxRabattTabelle(1).Left
    .Height = flxRabattTabelle(1).Height
    .Width = flxRabattTabelle(1).Width
        
    .FormatString = ">Schwellwert (EUR)|>Rabatt (%)"
    .SelectionMode = flexSelectionFree
    
    For i% = 0 To 1
        .ColWidth(i%) = (.Width - 90) / 2
    Next i%
    
'    .ColWidth(0) = TextWidth("Schwellwert (DM)     ")
'    .ColWidth(1) = TextWidth("Rabatt (%)    ")
'
'    spBreite% = 0
'    For i% = 0 To (.Cols - 1)
'        spBreite% = spBreite% + .ColWidth(i%)
'    Next i%
'    .Width = spBreite% + 90
    
    For i% = 0 To 4
        val1& = LifZus1.Schwellwert(i%)
        If (val1& > 0) Then .TextMatrix(i% + 1, 0) = Format(val1&, "0")
        dVal# = LifZus1.Rabatt(i%)
        If (dVal# > 0) Then .TextMatrix(i% + 1, 1) = Format(dVal#, "0.00")
    Next i%
End With

If (LifZus1.TabTyp) Then
    optRabattTabelle(1).Value = True
'    flxRabattTabelle(1).Visible = True
Else
    optRabattTabelle(0).Value = True
'    flxRabattTabelle(0).Visible = True
End If

cmdF5.Width = TextWidth(cmdF5.Caption) + 150
cmdF5.Height = wPara1.ButtonY
cmdF2.Width = cmdF5.Width
cmdF2.Height = wPara1.ButtonY

cmdF2.Top = flxRabattTabelle(0).Top
cmdF2.Left = flxRabattTabelle(0).Left + flxRabattTabelle(0).Width + 150
cmdF5.Top = cmdF2.Top + cmdF2.Height + 90
cmdF5.Left = cmdF2.Left

Hoehe2% = txtAusnahmenHerst(3).Top + txtAusnahmenHerst(3).Height + 2 * wPara1.TitelY
Breite2% = cmdF2.Left + cmdF2.Width + 2 * wPara1.LinksX
If (Hoehe2% > Hoehe%) Then Hoehe% = Hoehe2%
If (Breite2% > Breite%) Then Breite% = Breite2%

'--------------------------

tabStammdaten.Tab = 5
For i% = 0 To 6
    chkAusnahmen(i%).Value = LifZus1.AusnahmenKz(i%)
Next i%
chkAusnahmen(7).Value = LifZus1.AusnahmenAmRx
For i% = 0 To 9
    txtAusnahmenWg(i%).text = Format(LifZus1.AusnahmenWg(i%), "0")
Next i%
txtAusnahmenBM.text = Format(LifZus1.AusnahmenBM, "0")
txtAusnahmenZW.text = Format(LifZus1.AusnahmenZW, "0.00")
For i% = 0 To 5
    txtAusnahmenHerst(i%).text = LifZus1.AusnahmenHerst(i%)
Next i%

For i% = 0 To 7
    chkAusnahmen(i%).Left = wPara1.LinksX
Next i%
chkAusnahmen(0).Top = optRabattTabelle(0).Top
For i% = 1 To 7
    chkAusnahmen(i%).Top = chkAusnahmen(i% - 1).Top + chkAusnahmen(i% - 1).Height + 150
    If (i% = 4) Or (i% = 7) Then
        chkAusnahmen(i%).Top = chkAusnahmen(i%).Top + chkAusnahmen(i% - 1).Height + 150
    End If
Next i%

xpos% = chkAusnahmen(2).Left + chkAusnahmen(2).Width
ydiff% = (txtAusnahmenBM.Height - chkAusnahmen(0).Height) / 2

txtAusnahmenWg(0).Left = xpos%
txtAusnahmenWg(0).Top = chkAusnahmen(3).Top - ydiff%
For i% = 1 To 4
    txtAusnahmenWg(i%).Left = txtAusnahmenWg(i% - 1).Left + txtAusnahmenWg(i% - 1).Width + 45
    txtAusnahmenWg(i%).Top = txtAusnahmenWg(0).Top
Next i%
txtAusnahmenWg(5).Left = xpos%
txtAusnahmenWg(5).Top = txtAusnahmenWg(0).Top + txtAusnahmenWg(0).Height + 60
For i% = 6 To 9
    txtAusnahmenWg(i%).Left = txtAusnahmenWg(i% - 1).Left + txtAusnahmenWg(i% - 1).Width + 45
    txtAusnahmenWg(i%).Top = txtAusnahmenWg(5).Top
Next i%

txtAusnahmenBM.Left = xpos%
txtAusnahmenZW.Left = xpos%
txtAusnahmenBM.Top = chkAusnahmen(4).Top - ydiff%
txtAusnahmenZW.Top = chkAusnahmen(5).Top - ydiff%

txtAusnahmenHerst(0).Left = xpos%
txtAusnahmenHerst(0).Top = chkAusnahmen(6).Top - ydiff%
For i% = 1 To 2
    txtAusnahmenHerst(i%).Left = txtAusnahmenHerst(i% - 1).Left + txtAusnahmenHerst(i% - 1).Width + 45
    txtAusnahmenHerst(i%).Top = txtAusnahmenHerst(0).Top
Next i%
txtAusnahmenHerst(3).Left = xpos%
txtAusnahmenHerst(3).Top = txtAusnahmenHerst(0).Top + txtAusnahmenHerst(0).Height + 60
For i% = 4 To 5
    txtAusnahmenHerst(i%).Left = txtAusnahmenHerst(i% - 1).Left + txtAusnahmenHerst(i% - 1).Width + 45
    txtAusnahmenHerst(i%).Top = txtAusnahmenHerst(3).Top
Next i%

'--------------------------

tabStammdaten.Tab = 6
lblSonstiges(0).Left = wPara1.LinksX
lblSonstiges(0).Top = 3 * wPara1.TitelY   '2 * wPara1.TitelY

cboSonstiges(0).Left = lblSonstiges(0).Left + lblSonstiges(0).Width + 150
ydiff% = (cboSonstiges(0).Height - lblSonstiges(0).Height) / 2
cboSonstiges(0).Top = lblSonstiges(0).Top - ydiff%

If (LifZus1.WumsatzLief = 0) Then
    iLief%(0) = Val(StammdatenPzn$)
Else
    iLief%(0) = LifZus1.WumsatzLief
End If
If (LifZus1.MvdaLief = 0) Then
    iLief%(1) = Val(StammdatenPzn$)
Else
    iLief%(1) = LifZus1.MvdaLief
End If

cboSonstiges(0).Clear
cboSonstiges(1).Clear
For i% = 1 To Lif1.AnzRec
    Call Lif1.GetRecord(i% + 1)
    h$ = Lif1.kurz
    h$ = UCase(Trim$(h$))
    If (h$ <> "") Then
        If (Asc(Left$(h$, 1)) >= 32) Then
            h$ = h$ + " (" + Mid$(Str$(i%), 2) + ")"
            cboSonstiges(0).AddItem h$
            cboSonstiges(1).AddItem h$
            If (i% = iLief%(0)) Then h2$(0) = h$
            If (i% = iLief%(1)) Then h2$(1) = h$
        End If
    End If
Next i%
For j% = 0 To 1
    With cboSonstiges(j%)
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            If (.text = h2$(j%)) Then Exit For
        Next i%
    End With
Next j%
    
MaxWi% = 0
For i% = 1 To 2
    lblSonstiges(i%).Left = 2 * wPara1.LinksX
    wi% = lblSonstiges(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
txtSonstiges(0).Left = lblSonstiges(0).Left + MaxWi% + 300
cboSonstiges(1).Left = txtSonstiges(0).Left

lblSonstiges(1).Top = 2 * wPara1.TitelY + 60
lblSonstiges(2).Top = lblSonstiges(1).Top + lblSonstiges(1).Height + 150

ydiff% = (txtSonstiges(0).Height - lblSonstiges(1).Height) / Screen.TwipsPerPixelY
ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
txtSonstiges(0).Top = lblSonstiges(1).Top - ydiff%

ydiff% = (cboSonstiges(1).Height - lblSonstiges(2).Height) / 2
cboSonstiges(1).Top = lblSonstiges(2).Top - ydiff%

With chkSonstiges
    .Left = wPara1.LinksX
    .Top = lblSonstiges(0).Top + lblSonstiges(0).Height + 450
End With

With fmeSonstiges
    .Left = chkSonstiges.Left + chkSonstiges.Width + 150
    .Top = chkSonstiges.Top
    .Width = cboSonstiges(1).Left + cboSonstiges(1).Width + 2 * wPara1.LinksX
    .Height = cboSonstiges(1).Top + cboSonstiges(1).Height + wPara1.TitelY
End With

'--------------------------
MaxWi% = 0
For i% = 0 To 1
    lblSonstiges2(i%).Left = wPara1.LinksX + 150
    wi% = lblSonstiges2(i).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtSonstiges2(0).Left = lblSonstiges2(0).Left + MaxWi% + 300
txtSonstiges2(1).Left = txtSonstiges2(0).Left

lblSonstiges2(0).Top = 2 * wPara1.TitelY + 60
lblSonstiges2(1).Top = lblSonstiges2(0).Top + lblSonstiges2(0).Height + 150

txtSonstiges2(0).Top = lblSonstiges2(0).Top - ydiff%
txtSonstiges2(1).Top = lblSonstiges2(1).Top - ydiff%

With chkSonstiges2
    .Left = chkSonstiges.Left
    .Top = fmeSonstiges.Top + fmeSonstiges.Height + (300 * wPara1.BildFaktor)
End With

With fmeSonstiges2
    .Left = chkSonstiges2.Left + chkSonstiges2.Width + 150
    .Top = chkSonstiges2.Top
    .Width = txtSonstiges2(0).Left + txtSonstiges2(0).Width + 2 * wPara1.LinksX
    .Height = txtSonstiges2(1).Top + txtSonstiges2(1).Height + wPara1.TitelY
End With

'--------------------------

With fmeStammdaten(0)
    .Left = wPara1.LinksX
    .Top = 4 * wPara1.TitelY    '5
    .Width = Breite%
    .Height = Hoehe% - 30
    .Caption = ""
End With
For i% = 1 To 6
    With fmeStammdaten(i%)
        .Left = fmeStammdaten(0).Left
        .Top = fmeStammdaten(0).Top
        .Width = fmeStammdaten(0).Width
        .Height = fmeStammdaten(0).Height
        .Caption = ""
    End With
Next i%

With tabStammdaten
    .Left = wPara1.LinksX
    .Top = wPara1.TitelY
    .Width = fmeStammdaten(0).Left + fmeStammdaten(0).Width + 2 * wPara1.LinksX
    .Height = fmeStammdaten(0).Top + fmeStammdaten(0).Height + wPara1.TitelY
End With



'tabStammdaten.Height = Hoehe%
'tabStammdaten.Width = Breite% + wPara1.LinksX   'wegen 2.Zeile Tab

cmdOk.Top = tabStammdaten.Top + tabStammdaten.Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = tabStammdaten.Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdF5.Width = TextWidth(cmdF5.Caption) + 150
cmdF5.Height = cmdOk.Height
cmdF2.Width = cmdF5.Width
cmdF2.Height = cmdOk.Height

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
'    iAdd = wPara1.NlFlexBackY
'    iAdd2 = wPara1.NlCaptionY
'
'    With tabStammdaten
'        .Left = .Left + iAdd
'        .Top = .Top + iAdd
'    End With
'
'    cmdOk.Top = cmdOk.Top + 2 * iAdd
'    cmdEsc.Top = cmdOk.Top
'
'    Width = Width + 2 * iAdd
'    Height = Height + 2 * iAdd
'
'    tabStammdaten.Top = tabStammdaten.Top + iAdd2
'    cmdOk.Top = cmdOk.Top + iAdd2
'    cmdEsc.Top = cmdOk.Top
'    Height = Height + iAdd2
'
'    With nlcmdOk
'        .Init
'        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
'        .Top = tabStammdaten.Top + tabStammdaten.Height + iAdd + 600
'        .Caption = cmdOk.Caption
'        .TabIndex = cmdOk.TabIndex
'        .Enabled = cmdOk.Enabled
'        .default = cmdOk.default
'        .Cancel = cmdOk.Cancel
'        .Visible = True
'    End With
'    cmdOk.Visible = False
'
'    With nlcmdEsc
'        .Init
'        .Left = nlcmdOk.Left + .Width + 300
'        .Top = nlcmdOk.Top
'        .Caption = cmdEsc.Caption
'        .TabIndex = cmdEsc.TabIndex
'        .Enabled = cmdEsc.Enabled
'        .default = cmdEsc.default
'        .Cancel = cmdEsc.Cancel
'        .Visible = True
'    End With
'    cmdEsc.Visible = False
'
'    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450
'
'    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
''    RoundRect hdc, (tabStammdaten.Left - iAdd) / Screen.TwipsPerPixelX, (tabStammdaten.Top - iAdd) / Screen.TwipsPerPixelY, (tabStammdaten.Left + tabStammdaten.Width + iAdd) / Screen.TwipsPerPixelX, (tabStammdaten.Top + tabStammdaten.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'
''    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
''    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

Breite% = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
If (Breite% < 0) Then Breite% = 0
Me.Left = Breite%
Hoehe% = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
If (Hoehe% < 0) Then Hoehe% = 0
Me.Top = Hoehe%

If (InStr(Para1.Benutz, "w") = 0) Then
    chkStammdaten(0).Visible = False
    For i% = 0 To 2
        optStammdaten(i%).Visible = False
    Next i%
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Paint")
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
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (Para1.Newline) Then
'    iAdd = wPara1.NlFlexBackY
'    iAdd2 = wPara1.NlCaptionY
'
'    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
'    RoundRect hdc, (tabStammdaten.Left - iAdd) / Screen.TwipsPerPixelX, (tabStammdaten.Top - iAdd) / Screen.TwipsPerPixelY, (tabStammdaten.Left + tabStammdaten.Width + iAdd) / Screen.TwipsPerPixelX, (tabStammdaten.Top + tabStammdaten.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'
'    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

'Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, erg%, xpos%, ydiff%
'Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, maxSp%, val2%, iLief%
'Dim MenuHeight&, ScrollHeight&, val1&
'Dim val#
'Dim h$, h2$, FormStr$
'Dim c As Control
'
'Me.Width = tabStammdaten.Left + tabStammdaten.Width + 2 * wPara1.LinksX
'Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight
'Caption = AktWumsatzTyp$
'
'cmdOk.Left = (Me.ScaleWidth - (cmdOk.Width * 2 + 300)) / 2
'cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
'
'tabStammdaten.Tab = 0
'Call TabDisable
'Call TabEnable(0)

Private Sub tabStammdaten_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tabStammdaten_Click")
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
If (tabStammdaten.Visible = False) Then Call clsError.DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabStammdaten.Tab)

Call clsError.DefErrPop
End Sub

Sub StammdatenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StammdatenBefuellen")
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
Dim i%, j%, k%, l%, ind%
Dim h$, h2$, h3$

ind% = Val(StammdatenPzn$)
Lif1.GetRecord (ind% + 1)
LifZus1.GetRecord (ind% + 1)

For i% = 0 To 3
    txtStammdaten(i%).text = Trim(Lif1.Name(i%))
Next i%
txtStammdaten(4).text = Trim(Lif1.kurz)
txtStammdaten(5).text = Trim(Lif1.WarenBez)
txtStammdaten(6).text = Trim(Lif1.vertreter)
txtStammdaten(7).text = Trim(Lif1.telefon)
txtStammdaten(8).text = Trim(Lif1.Durchwahl)
txtStammdaten(9).text = Trim(Lif1.IdfApo)
txtStammdaten(10).text = Trim(Lif1.IdfLieferant)
txtStammdaten(11).text = LifZus1.ZusatzAufwand

chkStammdaten(0).Value = Abs(LifZus1.IstDirektLieferant)
fmeStammdaten3.Visible = chkStammdaten(0).Value

With chkStammdaten(1)
    .Value = Abs(LifZus1.IstAutoDirektLieferant)
    If (.Value) Then
        .ForeColor = chkStammdaten(0).ForeColor
    Else
        .ForeColor = vbRed
    End If
End With
optStammdaten(LifZus1.ZuordnungenModus).Value = True

txtStammdaten3(11).Visible = (StammdatenModus% = 1)

txtStammdaten2(0).text = Trim(Lif1.konto)
txtStammdaten2(1).text = Trim(Lif1.BLZ)
txtStammdaten2(2).text = Trim(Lif1.KontoVorschlag)
txtStammdaten2(3).text = Trim(Lif1.FibuKonto)
txtStammdaten2(4).text = ""
txtStammdaten2(5).text = Format(Lif1.WarenWert, "0.00")
txtStammdaten2(6).text = Format$(Lif1.skonto, "0.00")
txtStammdaten2(7).text = Format(Lif1.saldo, "0.00")
txtStammdaten2(8).text = LifZus1.Lieferfrist

chkStammdaten2(0).Value = Abs(Lif1.BankEinzug = "J")
chkStammdaten2(1).Value = Abs(LifZus1.DepotPflicht)
cboStammdaten2.ListIndex = InStr("DV", Lif1.DirektLieferant) - 1

txtStammdaten3(0).text = LifZus1.DirektMindestBM
txtStammdaten3(1).text = LifZus1.DirektMindestWert
txtStammdaten3(2).text = LifZus1.DirektMindestBewertung
txtStammdaten3(3).text = Format(LifZus1.KostenProLieferung, "0.00")
'txtStammdaten3(4).text = LifZus1.LiefFuerHerst(0)
'txtStammdaten3(5).text = LifZus1.Lieferfrist
txtStammdaten3(6).text = LifZus1.BevorratungsZeitraum
txtStammdaten3(7).text = LifZus1.ValutaStellung
txtStammdaten3(8).text = Format(LifZus1.FakturenRabatt, "0.00")
txtStammdaten3(9).text = LifZus1.TempBevorratungsZeitraum
txtStammdaten3(10).text = LifZus1.TempValutaStellung
txtStammdaten3(11).text = Format(LifZus1.TempFakturenRabatt, "0.00")
txtStammdaten3(12).text = Format(StammdatenWert, "0.00")
txtStammdaten3(13).text = LifZus1.LiefFuerHerst(1)
'txtStammdaten3(14).text = LifZus1.LiefFuerHerst(2)
txtStammdaten3(15).text = LifZus1.ProzentPlus

'chkStammdaten3.Value = Abs(LifZus1.DepotPflicht)
cboStammdaten3(0).ListIndex = LifZus1.FakturenRabattTyp
cboStammdaten3(1).ListIndex = LifZus1.PreisBasis

chkStammdaten4(0).Value = Abs(LifZus1.DirektBestModemKz)
chkStammdaten4(1).Value = Abs(LifZus1.DirektBestMailKz)
chkStammdaten4(2).Value = Abs(LifZus1.DirektBestFaxKz)
chkStammdaten4(3).Value = Abs(LifZus1.DirektBestDruckKz)
chkStammdaten4(4).Value = Abs(LifZus1.DirektBestComputerFaxKz)

txtStammdaten4(0).text = LifZus1.DirektBestModem
txtStammdaten4(1).text = LifZus1.DirektBestMail
txtStammdaten4(2).text = LifZus1.DirektBestFax
txtStammdaten4(3).text = LifZus1.AngebotWWW

chkSonstiges.Value = Abs(LifZus1.IstMvdaLieferant)
fmeSonstiges.Visible = chkSonstiges.Value
txtSonstiges(0).text = Format(LifZus1.MvdaProzent, "0.00")

chkSonstiges2.Value = Abs(LifZus1.HatFixRabatt)
fmeSonstiges2.Visible = chkSonstiges2.Value
For i = 0 To 1
    txtSonstiges2(i).text = Format(LifZus1.FixRabatt(i), "0.00")
Next i
    
OrgBevorratungsZeit% = Val(txtStammdaten3(9).text)
If (OrgBevorratungsZeit% = 0) Then OrgBevorratungsZeit% = Val(txtStammdaten3(6).text)
If (OrgBevorratungsZeit% = 0) Then OrgBevorratungsZeit% = Para1.BestellPeriode

OrgFakturenRabatt# = LifZus1.TempFakturenRabatt
If (OrgFakturenRabatt# = 0#) Then OrgFakturenRabatt# = LifZus1.FakturenRabatt

lstStammdaten3.Clear
For i% = 0 To 20
    h$ = Trim(LifZus1.LiefFuerHerst(i%))
    If (h$ <> "") Then lstStammdaten3.AddItem h$
Next i%
txtStammdaten3(4).text = ""

Call clsError.DefErrPop
End Sub

Sub StammdatenSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StammdatenSpeichern")
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
Dim i%, j%, k%, l%, ind%, iBevorratungsZeit%, AngebotePruefen%
Dim dVal#, iFakturenRabatt#
Dim h$, h2$, h3$

For i% = 0 To 3
    Lif1.Name(i%) = txtStammdaten(i%).text
Next i%
Lif1.kurz = txtStammdaten(4).text
Lif1.WarenBez = txtStammdaten(5).text
Lif1.vertreter = txtStammdaten(6).text
Lif1.telefon = txtStammdaten(7).text
Lif1.Durchwahl = txtStammdaten(8).text
Lif1.IdfApo = txtStammdaten(9).text
Lif1.IdfLieferant = txtStammdaten(10).text

dVal# = clsOpTool.xVal(txtStammdaten(11).text)
If (dVal# > 100#) Then dVal# = 100#
LifZus1.ZusatzAufwand = dVal#

'
Lif1.konto = txtStammdaten2(0).text
Lif1.BLZ = txtStammdaten2(1).text
Lif1.KontoVorschlag = txtStammdaten2(2).text
Lif1.FibuKonto = txtStammdaten2(3).text
Lif1.WarenWert = clsOpTool.xVal(txtStammdaten2(5).text)
Lif1.skonto = clsOpTool.xVal(txtStammdaten2(6).text)
Lif1.saldo = clsOpTool.xVal(txtStammdaten2(7).text)
LifZus1.Lieferfrist = clsOpTool.xVal(txtStammdaten2(8).text)

'
If (chkStammdaten2(0).Value) Then
    Lif1.BankEinzug = "J"
Else
    Lif1.BankEinzug = "N"
End If
'
LifZus1.DepotPflicht = chkStammdaten2(1).Value

ind% = cboStammdaten2.ListIndex
If (ind% = 0) Then
    Lif1.DirektLieferant = "D"
ElseIf (ind% = 1) Then
    Lif1.DirektLieferant = "V"
End If

LifZus1.IstDirektLieferant = chkStammdaten(0).Value
'LifZus1.IstAutoDirektLieferant = chkStammdaten(1).Value

For i% = 0 To 2
    If (optStammdaten(i%).Value) Then
        LifZus1.ZuordnungenModus = i%
        Exit For
    End If
Next i%

LifZus1.DirektMindestBM = clsOpTool.xVal(txtStammdaten3(0).text)
LifZus1.DirektMindestWert = clsOpTool.xVal(txtStammdaten3(1).text)
LifZus1.DirektMindestBewertung = clsOpTool.xVal(txtStammdaten3(2).text)
LifZus1.KostenProLieferung = clsOpTool.xVal(txtStammdaten3(3).text)
'LifZus1.LiefFuerHerst(0) = txtStammdaten3(4).text
'LifZus1.Lieferfrist = val(txtStammdaten3(5).text)
LifZus1.BevorratungsZeitraum = clsOpTool.xVal(txtStammdaten3(6).text)
LifZus1.ValutaStellung = clsOpTool.xVal(txtStammdaten3(7).text)
LifZus1.FakturenRabatt = clsOpTool.xVal(txtStammdaten3(8).text)

'LifZus1.DepotPflicht = chkStammdaten3.Value
LifZus1.FakturenRabattTyp = cboStammdaten3(0).ListIndex
LifZus1.PreisBasis = cboStammdaten3(1).ListIndex

LifZus1.DirektBestModemKz = chkStammdaten4(0).Value
LifZus1.DirektBestMailKz = chkStammdaten4(1).Value
LifZus1.DirektBestFaxKz = chkStammdaten4(2).Value
LifZus1.DirektBestDruckKz = chkStammdaten4(3).Value
LifZus1.DirektBestComputerFaxKz = chkStammdaten4(4).Value

LifZus1.DirektBestModem = txtStammdaten4(0).text
LifZus1.DirektBestMail = txtStammdaten4(1).text
LifZus1.DirektBestFax = txtStammdaten4(2).text
LifZus1.AngebotWWW = txtStammdaten4(3).text

LifZus1.TempBevorratungsZeitraum = clsOpTool.xVal(txtStammdaten3(9).text)
LifZus1.TempValutaStellung = clsOpTool.xVal(txtStammdaten3(10).text)
LifZus1.TempFakturenRabatt = clsOpTool.xVal(txtStammdaten3(11).text)

'LifZus1.LiefFuerHerst(1) = txtStammdaten3(13).text
'LifZus1.LiefFuerHerst(2) = txtStammdaten3(14).text

With lstStammdaten3
    For i% = 0 To 20
        If (i% < .ListCount) Then
            .ListIndex = i%
            LifZus1.LiefFuerHerst(i%) = .text
        Else
            LifZus1.LiefFuerHerst(i%) = ""
        End If
    Next i%
End With


k% = Val(txtStammdaten3(15).text)
If (k% > 255) Then k% = 255
LifZus1.ProzentPlus = k%

ind% = Val(StammdatenPzn$)
Lif1.PutRecord (ind% + 1)
LifZus1.PutRecord (ind% + 1)

If (StammdatenModus% = 1) Then
    AngebotePruefen% = False

    iBevorratungsZeit% = Val(txtStammdaten3(9).text)
    If (iBevorratungsZeit% = 0) Then iBevorratungsZeit% = Val(txtStammdaten3(6).text)
    If (iBevorratungsZeit% = 0) Then iBevorratungsZeit% = Para1.BestellPeriode
    
    If (iBevorratungsZeit <> OrgBevorratungsZeit%) Then
        Call StammdatenClass.CalcAlleDirektBM(iBevorratungsZeit%)
        AngebotePruefen% = True
    Else
        iFakturenRabatt# = LifZus1.TempFakturenRabatt
        If (iFakturenRabatt# = 0#) Then iFakturenRabatt# = LifZus1.FakturenRabatt
        If (iFakturenRabatt# <> OrgFakturenRabatt#) Then
            AngebotePruefen% = True
        End If
    End If
    If (AngebotePruefen%) Then
        Call StammdatenClass.CalcAlleDirektAngebote
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub tmrLiefStammdaten_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tmrLiefStammdaten_Timer")
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

tmrLiefStammdaten.Enabled = False
If (StammdatenStart$ = "") Then StammdatenStart$ = "0006"   ' "0103"
Call TabDisable
ind% = Val(Left$(StammdatenStart$, 2))
TxtInd% = Val(Mid$(StammdatenStart$, 3))
tabStammdaten.Tab = ind%
Call TabEnable(ind%)
TxtInd% = 0

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten_GotFocus")
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
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtStammdaten(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 1
If (index = 9) Or (index = 10) Then
    iEditModus% = 0
ElseIf (index = 11) Then
    iEditModus% = 4
End If

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten2_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten2_GotFocus")
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
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtStammdaten2(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 1
If (index = 2) Or (index = 3) Then
    iEditModus% = 0
ElseIf (index >= 5) And (index <= 7) Then
    iEditModus% = 4
End If

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten3_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten3_GotFocus")
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
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtStammdaten3(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Select Case index
    Case 4, 13, 14
        iEditModus% = 1
    Case 3, 8
        iEditModus% = 4
    Case Else
        iEditModus% = 0
End Select

If (index = 4) Then
    With cmdHersteller
        .Enabled = True
        .Caption = "H&inzufügen"
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten3_Change(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten3_Change")
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
Dim ind%, IstBewertungOk%, iBevorratungsZeit%
Dim dBewertung#

'If (StammdatenModus% = 1) And (index >= 9) And (txtStammdaten3(index).Visible) Then

If (StammdatenModus% = 1) And (txtStammdaten3(index).Visible) And (index <> 12) Then
    iBevorratungsZeit% = -1
    If (index = 6) Or (index = 9) Then
        iBevorratungsZeit% = Val(txtStammdaten3(9).text)
        If (iBevorratungsZeit% = 0) Then iBevorratungsZeit% = Val(txtStammdaten3(6).text)
        If (iBevorratungsZeit% = 0) Then iBevorratungsZeit% = Para1.BestellPeriode
    End If
    
    If (index >= 9) And (index <= 13) Then
        ind% = Val(StammdatenPzn$)
        LifZus1.GetRecord (ind% + 1)
        
        LifZus1.TempBevorratungsZeitraum = Val(txtStammdaten3(9).text)
        LifZus1.TempValutaStellung = clsOpTool.xVal(txtStammdaten3(10).text)
        LifZus1.TempFakturenRabatt = clsOpTool.xVal(txtStammdaten3(11).text)
    
    '    Call StammdatenSpeichern
        dBewertung# = StammdatenClass.CalcDirektBewertung#(IstBewertungOk%, False, True, False, iBevorratungsZeit%)
        txtStammdaten3(12).text = Format(dBewertung#, "0.0")
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten3_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten3_LostFocus")
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

If (index = 4) And (ActiveControl.Name <> cmdHersteller.Name) Then cmdHersteller.Enabled = False

Call clsError.DefErrPop
End Sub

Private Sub lstStammdaten3_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lstStammdaten3_GotFocus")
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

With cmdHersteller
    .Enabled = True
    .Caption = "&Entfernen"
End With

Call clsError.DefErrPop
End Sub

Private Sub lstStammdaten3_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lstStammdaten3_LostFocus")
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

If (ActiveControl.Name <> cmdHersteller.Name) Then cmdHersteller.Enabled = False

Call clsError.DefErrPop
End Sub

Private Sub txtStammdaten4_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtStammdaten4_GotFocus")
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
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtStammdaten4(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 1
If (index = 0) Or (index = 2) Then iEditModus% = 0

Call clsError.DefErrPop
End Sub

Private Sub chkAusnahmen_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkAusnahmen_Click")
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

If (index = 3) Then
    For i% = 0 To 9
        txtAusnahmenWg(i%).Enabled = chkAusnahmen(index).Value
    Next i%
ElseIf (index = 4) Then
    txtAusnahmenBM.Enabled = chkAusnahmen(index).Value
ElseIf (index = 5) Then
    txtAusnahmenZW.Enabled = chkAusnahmen(index).Value
ElseIf (index = 6) Then
    For i% = 0 To 5
        txtAusnahmenHerst(i%).Enabled = chkAusnahmen(index).Value
    Next i%
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF2_Click")
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
Dim i%, j%, ind%

ind% = 0
If (flxRabattTabelle(1).Visible) Then ind% = 1

With flxRabattTabelle(ind%)
    For j% = (.Rows - 2) To .row Step -1
        For i% = 0 To .Cols - 1
            .TextMatrix(j% + 1, i%) = .TextMatrix(j%, i%)
        Next i%
    Next j%
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

Call clsError.DefErrPop
End Sub

Private Sub cmdF5_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF5_Click")
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
Dim i%, ind%

ind% = 0
If (flxRabattTabelle(1).Visible) Then ind% = 1

With flxRabattTabelle(ind%)
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

Call clsError.DefErrPop
End Sub

Private Sub flxRabattTabelle_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxRabattTabelle_KeyDown")
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

If (KeyCode = vbKeyF2) Then
    cmdF2.Value = True
ElseIf (KeyCode = vbKeyF5) Then
    cmdF5.Value = True
End If

Call clsError.DefErrPop
End Sub

Sub EditSchwellwerteTxt()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSchwellwerteTxt")
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
Dim EditRow%, EditCol%, ind%, arow%
Dim dVal#
Dim h2$

ind% = ActiveControl.index

EditRow% = flxRabattTabelle(ind%).row
EditCol% = flxRabattTabelle(ind%).col

EditModus% = 0
If ((ind% = 0) And (EditCol% = 1)) Or ((ind% = 1) And (EditCol% = 4)) Then
    EditModus% = 4
End If

With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = True
    .row = arow%
End With
            
Load frmEdit2

With frmEdit2
    .Left = tabStammdaten.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Left = .Left + fmeStammdaten(4).Left
    .Top = tabStammdaten.Top + flxRabattTabelle(ind%).Top + EditRow% * flxRabattTabelle(ind%).RowHeight(0)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Top = .Top + fmeStammdaten(4).Top
    .Width = flxRabattTabelle(ind%).ColWidth(EditCol%)
    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit2.txtEdit
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = flxRabattTabelle(ind%).TextMatrix(EditRow%, EditCol%)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit2.Show 1
           
With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = False
    .row = arow%
            
    If (EditErg%) Then
        dVal# = Val(EditTxt$)
        If (EditModus% = 0) Then
            h2$ = Format(dVal#, "0")
        Else
            h2$ = Format(dVal#, "0.00")
        End If
        .TextMatrix(EditRow%, EditCol%) = h2$
        If (.col < .Cols - 2) Then .col = .col + 1
    End If
End With

Call clsError.DefErrPop
End Sub

Sub EditSchwellwerteLst()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSchwellwerteLst")
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
Dim i%, EditRow%, EditCol%, ind%, arow%
Dim dVal#
Dim h2$, s$

ind% = ActiveControl.index

EditRow% = flxRabattTabelle(ind%).row
EditCol% = flxRabattTabelle(ind%).col

With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = True
    .row = arow%
End With
            
Load frmEdit2

With frmEdit2
    .Left = tabStammdaten.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Left = .Left + fmeStammdaten(4).Left
    .Top = tabStammdaten.Top + flxRabattTabelle(ind%).Top + flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Top = .Top + fmeStammdaten(4).Top
    .Width = flxRabattTabelle(ind%).ColWidth(EditCol%)
    .Height = flxRabattTabelle(ind%).Height - flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
End With

With frmEdit2.lstEdit
    .Height = frmEdit2.ScaleHeight
    frmEdit2.Height = .Height
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    
    .Clear
    For i% = 0 To 4
        .AddItem VerglStr$(i%)
    Next i%
    
    .ListIndex = 0
    s$ = RTrim$(flxRabattTabelle(ind%).TextMatrix(EditRow%, EditCol%))
    If (s$ <> "") Then
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h2$ = .text
            If (s$ = h2$) Then
                Exit For
            End If
        Next i%
    End If
    
    .Visible = True
End With
   
frmEdit2.Show 1
           
With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = False
    .row = arow%
End With
            

If (EditErg%) Then
    h2$ = EditTxt$
    With flxRabattTabelle(ind%)
        .TextMatrix(EditRow%, EditCol%) = h2$
        If (.col < .Cols - 2) Then .col = .col + 1
    End With
End If

Call clsError.DefErrPop
End Sub

Sub EditSchwellwerteLstMulti()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSchwellwerteLstMulti")
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
Dim i%, EditRow%, EditCol%, ind%, arow%
Dim dVal#
Dim h2$, s$

ind% = ActiveControl.index

EditRow% = flxRabattTabelle(ind%).row
EditCol% = flxRabattTabelle(ind%).col

With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = True
    .row = arow%
End With
            
Load frmEdit2

With frmEdit2
    .Width = TextWidth("9 = AM RezPflicht mit Sonderregelung") + 3 * wPara1.FrmScrollHeight
'    .Left = tabStammdaten.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).ColPos(EditCol% - 2) + 45
    .Left = tabStammdaten.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).Width - .Width - 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Left = .Left + fmeStammdaten(4).Left
    .Top = tabStammdaten.Top + flxRabattTabelle(ind%).Top + flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Top = .Top + fmeStammdaten(4).Top
'    .Width = flxRabattTabelle(ind%).ColWidth(EditCol% - 2) + flxRabattTabelle(ind%).ColWidth(EditCol% - 1) + flxRabattTabelle(ind%).ColWidth(EditCol%)
'    .Height = flxRabattTabelle(ind%).Height - flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
    .Height = flxRabattTabelle(1).RowHeight(0) * 11
End With

With frmEdit2.lstMultiEdit
    .Height = frmEdit2.ScaleHeight
    frmEdit2.Height = .Height
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    
    .Clear
    .AddItem "(keine Gruupe)"
    .AddItem "0 = AM ApoPflicht"
    .AddItem "1 = AM RezPflicht"
    .AddItem "2 = AM Non ApoPflicht"
    .AddItem "3 = Nichtarzneimittel"
    .AddItem "4 = AM ApoPflicht, Non TaxPflicht"
    .AddItem "5 = AM RezPflicht, Non TaxPflicht"
    .AddItem "6 = AM RezPflicht + Rezepturzuschlag"
    .AddItem "7 = AM ApoPflicht + Rezepturzuschlag"
    .AddItem "8 = Droge, Chemikalien"
    .AddItem "9 = AM RezPflicht mit Sonderregelung"
    
    s$ = flxRabattTabelle(1).TextMatrix(EditRow%, EditCol%)
    For i% = 0 To 10
         .Selected(i%) = False
    Next i%
    For i% = 1 To 10
         h2$ = Format(i% - 1, "0")
         If (InStr(s$, h2$) > 0) Then
             .Selected(i%) = True
         End If
   Next i%
   .ListIndex = 0
    
    .Visible = True
End With
   
frmEdit2.Show 1
           
With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = False
    .row = arow%
End With
            

If (EditErg%) Then
    s$ = ""
    For i% = 0 To (EditAnzGefunden% - 1)
        h2$ = Format(EditGef%(i%) - 1, "0")
        s$ = s$ + h2$
    Next i%
    With flxRabattTabelle(ind%)
        .TextMatrix(EditRow%, EditCol%) = s$
    End With
End If

Call clsError.DefErrPop
End Sub

Sub SpeicherRabattTabelle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherRabattTabelle")
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
Dim i%, j%, k%, pos%, MitAusnahmen%, vergl%, ind%, val2%
Dim Schwell$, Rabatt$, h$, h2$

ind% = Val(StammdatenPzn$)
LifZus1.GetRecord (ind% + 1)

If (optRabattTabelle(0).Value) Then
    LifZus1.TabTyp = 0
Else
    LifZus1.TabTyp = 1
End If

With flxRabattTabelle(0)
    pos% = 0
    For i% = 1 To (.Rows - 1)
        Schwell$ = Trim(.TextMatrix(i%, 0))
        Rabatt$ = Trim(.TextMatrix(i%, 1))
        If (Rabatt$ <> "") Then
            LifZus1.Schwellwert(pos%) = Val(Schwell$)
            LifZus1.Rabatt(pos%) = CDbl(Rabatt$)
            pos% = pos% + 1
        End If
    Next i%
    Do
        If (pos% >= 5) Then Exit Do
        LifZus1.Schwellwert(pos%) = 0#
        LifZus1.Rabatt(pos%) = 0!
        pos% = pos% + 1
    Loop
End With

With flxRabattTabelle(1)
    pos% = 0
    For i% = 1 To (.Rows - 1)
        Rabatt$ = Trim(.TextMatrix(i%, 4))
        If (Rabatt$ <> "") Then
            LifZus1.AbBM(pos%) = Val(.TextMatrix(i%, 0))
            LifZus1.BisBM(pos%) = Val(.TextMatrix(i%, 1))
            
            vergl% = 0
            h$ = .TextMatrix(i%, 2)
            For j% = 0 To 4
                If (h$ = VerglStr$(j%)) Then
                    vergl% = j%
                    Exit For
                End If
            Next j%
            LifZus1.Vergleich(pos%) = vergl%
            
            LifZus1.zWert(pos%) = Val(.TextMatrix(i%, 3))
            LifZus1.BmEkRabatt(pos%) = CDbl(.TextMatrix(i%, 4))
            
        
            h$ = Trim(.TextMatrix(i%, 5))
            k% = 1
            val2% = 0
            For j% = 1 To 10
                h2$ = Format(j% - 1, "0")
                If (InStr(h$, h2$) > 0) Then
                    val2% = val2% + k%
                End If
                k% = k% * 2
            Next j%
            LifZus1.BmEkTabGruppe(pos%) = val2%
            
            pos% = pos% + 1
        End If
    Next i%
    Do
        If (pos% >= 5) Then Exit Do
        LifZus1.AbBM(pos%) = 0
        LifZus1.BisBM(pos%) = 0
        LifZus1.Vergleich(pos%) = 0
        LifZus1.zWert(pos%) = 0#
        LifZus1.BmEkRabatt(pos%) = 0#
        pos% = pos% + 1
    Loop
End With

MitAusnahmen% = False
For i% = 0 To 7
    If (i% = 7) Then
        LifZus1.AusnahmenAmRx = Abs(chkAusnahmen(i%).Value)
    Else
        LifZus1.AusnahmenKz(i%) = Abs(chkAusnahmen(i%).Value)
    End If
    If (chkAusnahmen(i%).Value) Then MitAusnahmen% = True
Next i%
LifZus1.HatAusnahmen = Abs(MitAusnahmen%)

For i% = 0 To 9
    LifZus1.AusnahmenWg(i%) = Val(txtAusnahmenWg(i%).text)
Next i%
LifZus1.AusnahmenBM = Val(txtAusnahmenBM.text)
LifZus1.AusnahmenZW = Val(txtAusnahmenZW.text)
For i% = 0 To 5
    LifZus1.AusnahmenHerst(i%) = txtAusnahmenHerst(i%).text
Next i%

For i% = 0 To 1
    h$ = LTrim$(RTrim$(cboSonstiges(i%).text))
    If (h$ <> "") Then
        ind% = InStr(h$, "(")
        If (ind% > 0) Then
            h$ = Mid$(h$, ind% + 1)
            ind% = InStr(h$, ")")
            h$ = Left$(h$, ind% - 1)
            If (i% = 0) Then
                LifZus1.WumsatzLief = Val(h$)
            Else
                LifZus1.MvdaLief = Val(h$)
            End If
        End If
    End If
Next i%

LifZus1.MvdaProzent = Val(txtSonstiges(0).text)
LifZus1.IstMvdaLieferant = chkSonstiges.Value

For i = 0 To 1
    LifZus1.FixRabatt(i) = Val(txtSonstiges2(i).text)
Next i
LifZus1.HatFixRabatt = chkSonstiges2.Value

ind% = Val(StammdatenPzn$)
LifZus1.PutRecord (ind% + 1)

Call clsError.DefErrPop
End Sub

Private Sub optRabattTabelle_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("optRabattTabelle_Click")
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

flxRabattTabelle(index).Visible = True
flxRabattTabelle((index + 1) Mod 2).Visible = False

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenWg_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenWg_GotFocus")
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

With txtAusnahmenWg(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 0

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenBM_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenBM_GotFocus")
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

With txtAusnahmenBM
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 0

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenZW_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenZW_GotFocus")
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
Dim h$

With txtAusnahmenZW
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 0

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenHerst_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenHerst_GotFocus")
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

With txtAusnahmenHerst(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 1

Call clsError.DefErrPop
End Sub

Private Sub txtSonstiges_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtSonstiges_GotFocus")
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

With txtSonstiges(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 4

Call clsError.DefErrPop
End Sub

Private Sub txtSonstiges2_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtSonstiges2_GotFocus")
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

With txtSonstiges2(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus% = 4

Call clsError.DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TabDisable")
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

For i% = 0 To 6
    fmeStammdaten(i%).Visible = False
Next i%
Call clsError.DefErrPop: Exit Sub

For i% = 0 To 10
    lblStammdaten(i%).Visible = False
    txtStammdaten(i%).Visible = False
    txtStammdaten3(i%).Visible = False
    If (i% < 10) Then lblStammdaten3(i%).Visible = False
    If (i% < 8) Then
        lblStammdaten2(i%).Visible = False
        txtStammdaten2(i%).Visible = False
    End If
    If (i% < 4) Then
        chkStammdaten4(i%).Visible = False
        txtStammdaten4(i%).Visible = False
    End If
Next i%
lblStammdaten4(0).Visible = False

chkStammdaten(0).Visible = False
chkStammdaten2(0).Visible = False
chkStammdaten2(1).Visible = False
cboStammdaten2.Visible = False
'chkStammdaten3.Visible = False
cboStammdaten3(0).Visible = False
cboStammdaten3(1).Visible = False
fmeStammdaten3.Visible = False

'''''''''''''''''''''''''''''''''''

For i% = 0 To 1
    flxRabattTabelle(i%).Visible = False
    optRabattTabelle(i%).Visible = False
Next i%
cmdF2.Visible = False
cmdF5.Visible = False


For i% = 0 To 7
    chkAusnahmen(i%).Visible = False
Next i%
For i% = 0 To 9
    txtAusnahmenWg(i%).Visible = False
Next i%
txtAusnahmenBM.Visible = False
txtAusnahmenZW.Visible = False
For i% = 0 To 5
    txtAusnahmenHerst(i%).Visible = False
Next i%

lblSonstiges(0).Visible = False
cboSonstiges(0).Visible = False

chkSonstiges.Visible = False
fmeSonstiges.Visible = False

chkSonstiges2.Visible = False
fmeSonstiges2.Visible = False

Call clsError.DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TabEnable")
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

fmeStammdaten(hTab%).Visible = True

If (hTab% = 0) Then
    For i% = 0 To 10
        lblStammdaten(i%).Visible = True
        txtStammdaten(i%).Visible = True
    Next i%
    If (InStr(Para1.Benutz, "w") = 0) Then
        chkStammdaten(0).Value = 0
    Else
        chkStammdaten(0).Visible = True
    End If
    If (Me.Visible) Then txtStammdaten(TxtInd%).SetFocus
ElseIf (hTab% = 1) Then
    For i% = 0 To 7
        lblStammdaten2(i%).Visible = True
        txtStammdaten2(i%).Visible = True
    Next i%
    txtStammdaten2(4).Visible = False
    chkStammdaten2(0).Visible = True
    chkStammdaten2(1).Visible = True
    cboStammdaten2.Visible = True
    If (Me.Visible) Then txtStammdaten2(TxtInd%).SetFocus
ElseIf (hTab% = 2) Then
    For i% = 0 To 12
        If (i% <> 5) Then
            If (i% < 10) Then lblStammdaten3(i%).Visible = True
            txtStammdaten3(i%).Visible = True
        End If
    Next i%
'    chkStammdaten3.Visible = True
    cboStammdaten3(0).Visible = True
    cboStammdaten3(1).Visible = True
    fmeStammdaten3.Visible = chkStammdaten(0).Value
    txtStammdaten3(12).Visible = (StammdatenModus% = 1)
    cmdHersteller.Enabled = False
    If (Me.Visible) Then txtStammdaten3(TxtInd%).SetFocus
ElseIf (hTab% = 3) Then
    For i% = 0 To 3
        chkStammdaten4(i%).Visible = True
        txtStammdaten4(i%).Visible = True
        If (i% < 3) Then
            txtStammdaten4(i%).Enabled = chkStammdaten4(i%).Value
        End If
    Next i%
    chkStammdaten4(4).Enabled = chkStammdaten4(2).Value
    lblStammdaten4(0).Visible = True
    If (Me.Visible) Then chkStammdaten4(TxtInd%).SetFocus
ElseIf (hTab% = 4) Then
    For i% = 0 To 1
'        flxRabattTabelle(i%).Visible = True
        optRabattTabelle(i%).Visible = True
    Next i%
    cmdF2.Visible = True
    cmdF5.Visible = True
    i% = 1
    If (optRabattTabelle(0).Value) Then i% = 0
    flxRabattTabelle(i%).Visible = True
    optRabattTabelle(i%).SetFocus
ElseIf (hTab% = 5) Then
    For i% = 0 To 7
        chkAusnahmen(i%).Visible = True
    Next i%
    For i% = 0 To 9
        txtAusnahmenWg(i%).Visible = True
    Next i%
    txtAusnahmenBM.Visible = True
    txtAusnahmenZW.Visible = True
    For i% = 0 To 5
        txtAusnahmenHerst(i%).Visible = True
    Next i%
    chkAusnahmen(0).SetFocus
Else
    lblSonstiges(0).Visible = True
    cboSonstiges(0).Visible = True
    cboSonstiges(0).SetFocus

    chkSonstiges.Visible = True
    fmeSonstiges.Visible = chkSonstiges.Value
    chkSonstiges2.Visible = True
    fmeSonstiges2.Visible = chkSonstiges2.Value
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseDown")
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
    
If (y <= wPara1.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseMove")
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
Dim c As Object

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is nlCommand) Then
        If (c.MouseOver) Then
            c.MouseOver = 0
        End If
    End If
Next
On Error GoTo DefErr

Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
If (iNewLine) And (Me.Visible) Then
'    CurrentX = wPara1.NlFlexBackY
'    CurrentY = (wPara1.NlCaptionY - TextHeight(Caption)) / 2
'    ForeColor = vbBlack
'    Me.Print Caption
End If
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
'    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
'        Exit Sub
'    ElseIf (KeyAscii = 27) Then
'        Call nlcmdEsc_Click
'        Exit Sub
'    End If
End If

If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((iEditModus% <> 4) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub



