VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmnlOptionen 
   Caption         =   "Optionen"
   ClientHeight    =   6975
   ClientLeft      =   -1425
   ClientTop       =   330
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdChange 
      Caption         =   "Hinunter (>)"
      Height          =   450
      Index           =   1
      Left            =   8400
      TabIndex        =   26
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Hinauf (<)"
      Height          =   450
      Index           =   0
      Left            =   8400
      TabIndex        =   25
      Top             =   6000
      Width           =   1200
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "Entfernen (F5)"
      Height          =   450
      Left            =   7080
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Einfügen (F2)"
      Height          =   450
      Left            =   5760
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox lstOptionenMulti 
      Height          =   450
      Left            =   3600
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtOptionen 
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstOptionen 
      Height          =   450
      Left            =   6000
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2280
      TabIndex        =   11
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   10
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   5325
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9393
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   6
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
      TabCaption(0)   =   "&1 - Kontrollen"
      TabPicture(0)   =   "nlOptionen.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxOptionen(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Zuordnungen"
      TabPicture(1)   =   "nlOptionen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxOptionen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Rufzeiten"
      TabPicture(2)   =   "nlOptionen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxOptionen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Automatik"
      TabPicture(3)   =   "nlOptionen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblOptionenAutomatik(1)"
      Tab(3).Control(1)=   "lblOptionenAutomatik(0)"
      Tab(3).Control(2)=   "lblOptionenAutomatikMinuten(0)"
      Tab(3).Control(3)=   "lblOptionenAutomatikMinuten(1)"
      Tab(3).Control(4)=   "lblOptionenAutomatikMinuten(2)"
      Tab(3).Control(5)=   "lblOptionenAutomatik(2)"
      Tab(3).Control(6)=   "fmeOptionenAutomatikBestVors"
      Tab(3).Control(7)=   "txtOptionenAutomatik(1)"
      Tab(3).Control(8)=   "txtOptionenAutomatik(0)"
      Tab(3).Control(9)=   "txtOptionenAutomatik(2)"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "&5 - Schwellwerte"
      TabPicture(4)   =   "nlOptionen.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeSchwellwerte"
      Tab(4).Control(1)=   "chkSchwellwerteAktiv"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "&6 - Spez.Ausw."
      TabPicture(5)   =   "nlOptionen.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "flxAbsagenKz"
      Tab(5).Control(1)=   "flxAllgLieferanten"
      Tab(5).Control(2)=   "flxFeiertage"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "&7 - Direktbez."
      TabPicture(6)   =   "nlOptionen.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fmeOptionenDirektBezug"
      Tab(6).Control(1)=   "fmeOptionenDirektBezug2"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "&8 - Automatenware"
      TabPicture(7)   =   "nlOptionen.frx":00C4
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "lblAutomatenWare"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "flxAutomatenWare"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "txtAutomatenWare"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      Begin VB.TextBox txtAutomatenWare 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         MaxLength       =   1
         TabIndex        =   54
         Text            =   "XX"
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Frame fmeOptionenDirektBezug2 
         Caption         =   "Allgemein"
         Height          =   1575
         Left            =   -73080
         TabIndex        =   49
         Top             =   3360
         Width           =   7335
         Begin VB.CheckBox chkOptionenDirektBezug2 
            Caption         =   "Artikel mit &BM=0 in Komplett-Übersicht anzeigen"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   52
            Top             =   1200
            Width           =   6735
         End
         Begin VB.CheckBox chkOptionenDirektBezug2 
            Caption         =   "auf Vorhandensein in &WÜ prüfen"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   51
            Top             =   840
            Width           =   6735
         End
         Begin VB.CheckBox chkOptionenDirektBezug2 
            Caption         =   "akt. &Lagerstände berücksichtigen"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   50
            Top             =   360
            Width           =   6735
         End
      End
      Begin VB.Frame fmeOptionenDirektBezug 
         Caption         =   "Direktbezugs-Automatik"
         Height          =   2295
         Left            =   -73200
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtOptionenDirektBezug 
            Alignment       =   1  'Rechts
            Height          =   375
            Index           =   0
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   44
            Text            =   "9999"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtOptionenDirektBezug 
            Alignment       =   1  'Rechts
            Height          =   375
            Index           =   1
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   46
            Text            =   "9999"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtOptionenDirektBezug 
            Alignment       =   1  'Rechts
            Height          =   375
            Index           =   2
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   48
            Text            =   "9999"
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblOptionenDirektBezug2 
            Caption         =   "Minuten"
            Height          =   255
            Index           =   2
            Left            =   5400
            TabIndex        =   58
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblOptionenDirektBezug2 
            Caption         =   "Minuten"
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   57
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblOptionenDirektBezug2 
            Caption         =   "Minuten"
            Height          =   255
            Index           =   0
            Left            =   5400
            TabIndex        =   56
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblOptionenDirektBezug 
            Caption         =   "Automat. Sendung nur durchführen, wenn nächste &Rufzeit in mind."
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label lblOptionenDirektBezug 
            Caption         =   "Automat. Sendung von Aufträgen mit &Kontrollen nach"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label lblOptionenDirektBezug 
            Caption         =   "&Hinweis zur Kontrolle alle"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   3615
         End
      End
      Begin VB.Frame fmeSchwellwerte 
         Caption         =   "Schwellwert-Parameter"
         Height          =   3135
         Left            =   -74400
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CheckBox chkSchwellwerte 
            Caption         =   "&Beobachtungszeitraum glätten"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   2760
            Width           =   5535
         End
         Begin VB.TextBox txtSchwellwerte 
            Height          =   375
            Index           =   3
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   36
            Text            =   "999"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtSchwellwerte 
            Height          =   375
            Index           =   2
            Left            =   4320
            MaxLength       =   3
            TabIndex        =   35
            Text            =   "999"
            Top             =   1200
            Width           =   735
         End
         Begin VB.CheckBox chkSchwellwerte 
            Caption         =   "Berechnung auch bei Sende&hinweis durchführen"
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   38
            Top             =   2400
            Width           =   5535
         End
         Begin VB.TextBox txtSchwellwerte 
            Height          =   375
            Index           =   1
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "999"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtSchwellwerte 
            Height          =   375
            Index           =   0
            Left            =   4440
            MaxLength       =   3
            TabIndex        =   33
            Text            =   "999"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblSchwellwerte 
            Caption         =   "Warnung ab x Sendungen vor Monatsende"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   3615
         End
         Begin VB.Label lblSchwellwerte 
            Caption         =   "&Warnung bei Unterschreiten Schwellwert (%)"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   39
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblSchwellwerte 
            Caption         =   "&Toleranz zum Erreichen des Schwellwertes (%)"
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label lblSchwellwerte 
            Caption         =   "&Zeitfenster der berücksichtigten Lieferanten (Minuten)"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.CheckBox chkSchwellwerteAktiv 
         Caption         =   "Schwellwert-&Automatik aktiv"
         Height          =   375
         Left            =   -74280
         TabIndex        =   32
         Top             =   1320
         Width           =   5535
      End
      Begin VB.TextBox txtOptionenAutomatik 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -67320
         TabIndex        =   5
         Text            =   "999"
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtOptionenAutomatik 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -67080
         TabIndex        =   3
         Text            =   "999"
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtOptionenAutomatik 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -67080
         TabIndex        =   4
         Text            =   "999"
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   2400
         Width           =   495
      End
      Begin VB.Frame fmeOptionenAutomatikBestVors 
         Caption         =   "Bestell&vorschlag"
         Height          =   2055
         Left            =   -74640
         TabIndex        =   15
         Top             =   3720
         Width           =   9015
         Begin VB.CheckBox chkOptionenAutomatikBestVors 
            Caption         =   "&Komplett vor jedem eingetragenen Sendeauftrag :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtOptionenAutomatikBestVors 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   7
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkOptionenAutomatikBestVors 
            Caption         =   "&Periodisch im Hintergrund, ein Durchlauf dauert"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox txtOptionenAutomatikBestVors 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   9
            Text            =   "999"
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblOptionenAutomatikBestVorsMinuten 
            Caption         =   "Minuten davor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   7200
            TabIndex        =   17
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblOptionenAutomatikBestVorsMinuten 
            Caption         =   "Minuten "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   7200
            TabIndex        =   16
            Top             =   1080
            Width           =   1695
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   0
         Left            =   -74640
         TabIndex        =   0
         Top             =   1800
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   1
         Left            =   -74760
         TabIndex        =   1
         Top             =   1620
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   2
         Left            =   -74760
         TabIndex        =   2
         Top             =   1740
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxFeiertage 
         Height          =   1680
         Left            =   -73800
         TabIndex        =   37
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2963
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
      Begin MSFlexGridLib.MSFlexGrid flxAllgLieferanten 
         Height          =   1680
         Left            =   -71160
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2963
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
      Begin MSFlexGridLib.MSFlexGrid flxAbsagenKz 
         Height          =   1680
         Left            =   -68640
         TabIndex        =   60
         Top             =   1680
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2963
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
      Begin MSFlexGridLib.MSFlexGrid flxAutomatenWare 
         Height          =   2040
         Left            =   1440
         TabIndex        =   55
         Top             =   2160
         Width           =   5295
         _ExtentX        =   9340
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
      End
      Begin VB.Label lblAutomatenWare 
         Caption         =   "&Lagercode der zu prüfenden Angebotszuordnungen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   53
         Top             =   1080
         Width           =   6855
      End
      Begin VB.Label lblOptionenAutomatik 
         Caption         =   "Maximale Verspätung für Bereitstellung :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74760
         TabIndex        =   28
         Top             =   3000
         Width           =   6855
      End
      Begin VB.Label lblOptionenAutomatikMinuten 
         Caption         =   "Minuten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -66840
         TabIndex        =   27
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblOptionenAutomatikMinuten 
         Caption         =   "Minuten "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -66600
         TabIndex        =   21
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblOptionenAutomatikMinuten 
         Caption         =   "Minuten davor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -66600
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblOptionenAutomatik 
         Caption         =   "Wieviele Minuten vor dem &Senden soll der Hinweis erscheinen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74520
         TabIndex        =   19
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Label lblOptionenAutomatik 
         Caption         =   "Wieviele Minuten soll auf &Anruf gewartet werden :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74520
         TabIndex        =   18
         Top             =   2400
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmnlOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "NLOPTIONEN.FRM"

Private Sub cmdChange_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdChange_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, row%
Dim h$
        
With flxOptionen(tabOptionen.Tab)
    .redraw = False
    row% = .row
    If (Index = 0) Then
        If (row% > 1) Then
            For i% = 0 To .Cols - 1
                h$ = .TextMatrix(row% - 1, i%)
                .TextMatrix(row% - 1, i%) = .TextMatrix(row%, i%)
                .TextMatrix(row%, i%) = h$
            Next i%
            .row = row% - 1
        End If
    Else
        If (row% < (.Rows - 1)) Then
            For i% = 0 To .Cols - 1
                h$ = .TextMatrix(row% + 1, i%)
                .TextMatrix(row% + 1, i%) = .TextMatrix(row%, i%)
                .TextMatrix(row%, i%) = h$
            Next i%
            .row = row% + 1
        End If
    End If
    .redraw = True
End With

Call DefErrPop

End Sub

Private Sub cmdEsc_Click()
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

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF2_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
        
With flxOptionen(tabOptionen.Tab)
    For j% = (.Rows - 2) To .row Step -1
        For i% = 0 To .Cols - 1
            .TextMatrix(j% + 1, i%) = .TextMatrix(j%, i%)
        Next i%
    Next j%
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

Call DefErrPop
End Sub

Private Sub cmdF5_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF5_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

With flxOptionen(tabOptionen.Tab)
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

Call DefErrPop
End Sub

Private Sub cmdOk_Click()
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
Dim i%, j%, l%, hTab%, row%, col%, Anzgefunden%, ind%, gef%(19)
Dim uhr%, st%, min%, MultiAuswahl%, LiefGef%, BmMulti%
Dim h$, h2$, s$, BetrLief$, Lief2$

If (ActiveControl.Name = cmdOk.Name) Then
    Call AuslesenFlexKontrollen
    Call SpeicherIniKontrollen
    Call SpeicherIniZuordnungen
    Call SpeicherIniRufzeiten
    Call HoleRufzeitenLieferanten
    Call SpeicherIniFeiertage
    Call frmAction.SpeicherIniWerte
    
    Call ResetZuKontrollieren
    OptionenNeu% = True
    Unload Me

Else
    hTab% = tabOptionen.Tab
    If (hTab% = 7) Then
        If (ActiveControl.Name = txtAutomatenWare.Name) Then
            SendKeys "{TAB}", True
        Else
            Call EditAutomatenLiefs
        End If
    ElseIf (hTab% < 3) Then
        row% = flxOptionen(hTab%).row
        col% = flxOptionen(hTab%).col
        If (lstOptionen.Visible = True) Then
            h$ = lstOptionen.text
            lstOptionen.Visible = False
            
            If ((h$ = "Eingabe") Or (h$ = "?-fach")) Then
                With txtOptionen
                    BmMulti% = False
                    If (h$ = "?-fach") Then BmMulti% = True
                    
                    .Top = tabOptionen.Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
                    .Height = flxOptionen(hTab%).RowHeight(1)
                    .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 15
                    .Width = flxOptionen(hTab%).ColWidth(col%)
                    
                    h2$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    If (BmMulti%) Then
                        ind% = InStr(h2$, "*")
                        If (ind% > 0) Then
                            .text = Left$(h2$, ind% - 1)
                        Else
                            .text = h2$
                        End If
                    Else
                        ind% = InStr(h2$, "Tage")
                        If (ind% > 0) Then
                            .text = Left$(h2$, ind% - 1)
                        Else
                            .text = h2$
                        End If
                    End If
                    .SelStart = 0
                    .SelLength = Len(.text)
                    .BackColor = vbRed
                    .Visible = True
                    .TabStop = True
                    .SetFocus
                End With
            End If
        ElseIf (txtOptionen.Visible = True) Then
            h$ = RTrim(txtOptionen.text)
            txtOptionen.Visible = False
            tabOptionen.Enabled = True
            With flxOptionen(hTab%)
                If (hTab% < 2) Then
                    If (BmMulti% = True) Then
                        .TextMatrix(.row, .col) = h$ + "* BMopt"
                    Else
                        If (col% = 2) And (.TextMatrix(row%, 0) = "Besorger") Then h$ = h$ + " Tage"
                        .TextMatrix(.row, .col) = h$
                    End If
                Else
                    Select Case .col
                        Case 1
                            uhr% = Val(h$)
                            st% = uhr% \ 100
                            min% = uhr% Mod 100
                            If (st% >= 0) And (st% < 24) And (min% >= 0) And (min% < 60) Then
                                .TextMatrix(.row, .col) = Format(st%, "00") + ":" + Format(min%, "00")
                            End If
                        Case 4, 5
                            .TextMatrix(.row, .col) = UCase$(h$)
                    End Select
                End If
                .SetFocus
                If (.col < .Cols - 2) Then .col = .col + 1
                If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
            End With
        Else
            s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
            
            Select Case flxOptionen(hTab%).col
                Case 0
                    If (hTab% <> 3) Then
                        Call EditOptionenLst
                    End If
                    Call DefErrPop: Exit Sub
                Case 1
                    If (hTab% <> 2) Then
                        Call EditOptionenLst
                    Else
                        Call EditOptionenTxt
                    End If
                Case 2
                    If (hTab% < 2) Then
                        If (ZeilenTyp%(s$) = 1) Then
                            If (UCase(s$) = "LIEFERANT") Then
                                Call EditOptionenLst
                            Else
                                Call EditOptionenTxt
                            End If
                        Else
                            Call EditOptionenTxt
                        End If
                    Else
                        Call EditOptionenTxt
                    End If
                Case 3
                    If (hTab% = 0) Then
                        Call EditOptionenLst
                    Else
                        Call EditOptionenLstMulti
                    End If
                Case 5, 6
                    Call EditOptionenLst
                Case 7
                    If (hTab% = 2) Then
                        Call EditOptionenLst
                    End If
            End Select
        End If
    End If
End If

Call DefErrPop
End Sub

Private Sub flxOptionen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen_KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Index < 3) Then
    If (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) Then
        cmdF5.Value = True
    End If
End If
Call DefErrPop
End Sub

Private Sub flxOptionen_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (KeyAscii = Asc("<")) Then
    cmdChange(0).Value = True
ElseIf (KeyAscii = Asc(">")) Then
    cmdChange(1).Value = True
End If

Call DefErrPop

End Sub

Private Sub flxFeiertage_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxFeiertage_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (KeyAscii = Asc(" ")) Then
    With flxFeiertage
        If (.TextMatrix(.row, 0) = Chr$(214)) Then
            .TextMatrix(.row, 0) = " "
        Else
            .TextMatrix(.row, 0) = Chr$(214)
        End If
    End With
End If

Call DefErrPop

End Sub

Private Sub flxAllgLieferanten_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAllgLieferanten_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (KeyAscii = Asc(" ")) Then
    With flxAllgLieferanten
        If (.TextMatrix(.row, 0) = Chr$(214)) Then
            .TextMatrix(.row, 0) = " "
        Else
            .TextMatrix(.row, 0) = Chr$(214)
        End If
    End With
End If

Call DefErrPop

End Sub

Private Sub flxAbsagenKz_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAbsagenKz_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (KeyAscii = Asc(" ")) Then
    With flxAbsagenKz
        If (.TextMatrix(.row, 0) = Chr$(214)) Then
            .TextMatrix(.row, 0) = " "
        Else
            .TextMatrix(.row, 0) = Chr$(214)
        End If
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, Hoehe3%, xpos%, ydiff%
Dim h$, h2$, FormStr$
Dim c As Control

Call wpara.InitFont(Me)

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr

tabOptionen.Left = wpara.LinksX
tabOptionen.Top = wpara.TitelY

For j% = 0 To 2
    With flxOptionen(j%)
        If (j% = 2) Then
            .Cols = 9
        Else
            .Cols = 5
        End If
        
        .Rows = 2
        .FixedRows = 1
        
        .Top = 900   'TitelY%
        .Left = wpara.LinksX
        .Height = .RowHeight(0) * 11 + 90
        
        If (j% = 2) Then
            FormStr$ = "^Lieferant|^Rufzeit|^Liefzeit|^Wochentag(e)||^Erg|^Art|<Aktiv|"
        Else
            FormStr$ = "^Wert1|^Bedingung|^Wert2|"
            If (j% = 0) Then
                FormStr$ = FormStr$ + "<Unkontrollierte"
            Else
                FormStr$ = FormStr$ + "<Lieferant(en)"
            End If
            FormStr$ = FormStr$ + "|^ "
        End If
        
        .FormatString = FormStr$
        .Rows = 1
        .SelectionMode = flexSelectionFree
    End With
Next j%
        


Font.Bold = False   ' True

tabOptionen.Tab = 2
With flxOptionen(2)
    .ColWidth(0) = TextWidth("WWWWWW (999)")
    .ColWidth(1) = TextWidth("Rufzeit  ") 'WWWWWW
    .ColWidth(2) = TextWidth("Liefzeit  ") 'WWWWWW
'    .ColWidth(2) = TextWidth(String$(15, "W"))
    .ColWidth(3) = TextWidth("Wwwwwwwww, Wwwwwwwww")
    .ColWidth(4) = 0
    .ColWidth(5) = TextWidth("WWW")
    .ColWidth(6) = TextWidth("WWW")
    .ColWidth(7) = TextWidth("WWWWW")
    .ColWidth(8) = 0
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

For j% = 0 To 1
    tabOptionen.Tab = j%
    With flxOptionen(j%)
        .Width = flxOptionen(2).Width
        .ColWidth(0) = 0
        .ColWidth(4) = 0
        For i% = 1 To (.Cols - 2)
            .ColWidth(i%) = .Width / 4
        Next i%

        spBreite% = 0
        For i% = 1 To .Cols - 2
            spBreite% = spBreite% + .ColWidth(i%)
        Next i%
        .ColWidth(0) = .Width - spBreite%
    End With
Next j%

'------------------------
tabOptionen.Tab = 4

With chkSchwellwerteAktiv
    .Top = 900
    .Left = wpara.LinksX
End With

For i% = 0 To 3
    lblSchwellwerte(i%).Left = wpara.LinksX
Next i%
lblSchwellwerte(0).Top = 2 * wpara.TitelY + 60
For i% = 1 To 3
    lblSchwellwerte(i%).Top = lblSchwellwerte(i% - 1).Top + lblSchwellwerte(i% - 1).Height + 150
Next i%

xpos% = lblSchwellwerte(0).Left + lblSchwellwerte(0).Width
ydiff% = (txtSchwellwerte(0).Height - lblSchwellwerte(0).Height) / 2

For i% = 0 To 3
    txtSchwellwerte(i%).Left = xpos%
    txtSchwellwerte(i%).Top = lblSchwellwerte(i%).Top - ydiff%
Next i%

For i% = 0 To 1
    chkSchwellwerte(i%).Left = lblSchwellwerte(0).Left
Next i%
chkSchwellwerte(0).Top = lblSchwellwerte(3).Top + lblSchwellwerte(3).Height + 300
For i% = 1 To 1
    chkSchwellwerte(i%).Top = chkSchwellwerte(i% - 1).Top + chkSchwellwerte(i% - 1).Height + 90
Next i%

'fmeSchwellwerte.Left = wpara.LinksX
'fmeSchwellwerte.Top = chkSchwellwerteAktiv.Top + chkSchwellwerteAktiv.Height + 300
fmeSchwellwerte.Left = chkSchwellwerteAktiv.Left + chkSchwellwerteAktiv.Width + 300
fmeSchwellwerte.Top = chkSchwellwerteAktiv.Top

fmeSchwellwerte.Width = txtSchwellwerte(2).Left + txtSchwellwerte(2).Width + 2 * wpara.LinksX
'fmeSchwellwerte.Height = txtSchwellwerte(2).Top + txtSchwellwerte(2).Height + 2 * wpara.TitelY
fmeSchwellwerte.Height = chkSchwellwerte(1).Top + chkSchwellwerte(1).Height + wpara.TitelY


'------------------------
tabOptionen.Tab = 5

With flxFeiertage
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    
    .Top = 900   'TitelY%
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    FormStr$ = "|<Tag|<Datum"
    .FormatString = FormStr$
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = TextWidth("X  ")
    .ColWidth(1) = TextWidth("Tag der deutschen Einheitww")
    .ColWidth(2) = TextWidth("99.99.9999www")
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

With flxAllgLieferanten
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    
    .Top = 900   'TitelY%
    .Left = flxFeiertage.Left + flxFeiertage.Width + 300
    .Height = .RowHeight(0) * 11 + 90
    
    FormStr$ = "|<Lief. für allg. Angebote"
    .FormatString = FormStr$
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = TextWidth("X  ")
    .ColWidth(1) = TextWidth("Lief. für allg. Angebotew")
    .ColWidth(2) = 0
    
    spBreite% = 0
    For i% = 0 To .Cols - 2
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With


With flxAbsagenKz
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    
    .Top = 900   'TitelY%
    .Left = flxAllgLieferanten.Left + flxAllgLieferanten.Width + 300
    .Height = .RowHeight(0) * 11 + 90
    
    FormStr$ = "|<Absagen mit NL"
    .FormatString = FormStr$
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = TextWidth("X  ")
    .ColWidth(1) = TextWidth("Absagen mit NLwww")
    .ColWidth(2) = 0
    
    spBreite% = 0
    For i% = 0 To .Cols - 2
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .Rows = 1
    
    If (para.Land <> "A") Then
        .Visible = False
    End If
End With


'------------------------
tabOptionen.Tab = 7
    
With lblAutomatenWare
    .Left = wpara.LinksX
    .Top = 900
    xpos% = .Left + .Width
End With
With txtAutomatenWare
    .Left = xpos%
    .Top = lblAutomatenWare.Top - ydiff%
End With
With flxAutomatenWare
    .Cols = 3
    .Rows = 11
    .FixedRows = 1
    
    .Top = lblAutomatenWare.Top + lblAutomatenWare.Height + 300
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 6 + 90
    
    FormStr$ = "^zugeordneter Lief.|^wird ersetzt durch|"
    .FormatString = FormStr$
    .SelectionMode = flexSelectionFree
    
    .ColWidth(0) = TextWidth(String(25, "X"))
    .ColWidth(1) = TextWidth(String(25, "X"))
    .ColWidth(2) = 0
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
End With


'------------------------

Call OptionenBefuellen
'

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdF2.Width = TextWidth(cmdF5.Caption) + 150
cmdF2.Height = wpara.ButtonY
cmdF2.Left = tabOptionen.Left + flxOptionen(0).Left + flxOptionen(0).Width + 150
cmdF2.Top = tabOptionen.Top + flxOptionen(0).Top

cmdF5.Width = cmdF2.Width
cmdF5.Height = cmdF2.Height
cmdF5.Left = cmdF2.Left
cmdF5.Top = cmdF2.Top + cmdF2.Height + 90

For i% = 0 To 1
    With cmdChange(i%)
        .Width = cmdF2.Width
        .Height = cmdF2.Height
        .Left = cmdF2.Left
    End With
Next i%
cmdChange(1).Top = tabOptionen.Top + flxOptionen(0).Top + flxOptionen(0).Height - cmdChange(1).Height
cmdChange(0).Top = cmdChange(1).Top - cmdChange(0).Height - 90


Hoehe1% = flxOptionen(2).Top + flxOptionen(2).Height + 180
Breite1% = cmdF2.Left + cmdF2.Width + wpara.LinksX - tabOptionen.Left

'------------------------



tabOptionen.Tab = 3


txtOptionenAutomatik(0).Top = 900   'TitelY%
For i% = 1 To 2
    txtOptionenAutomatik(i%).Top = txtOptionenAutomatik(i% - 1).Top + txtOptionenAutomatik(i% - 1).Height + 90
Next i%

lblOptionenAutomatik(0).Left = wpara.LinksX
lblOptionenAutomatik(0).Top = txtOptionenAutomatik(0).Top
For i% = 1 To 2
    lblOptionenAutomatik(i%).Left = lblOptionenAutomatik(i% - 1).Left
    lblOptionenAutomatik(i%).Top = txtOptionenAutomatik(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblOptionenAutomatik(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtOptionenAutomatik(0).Left = lblOptionenAutomatik(0).Left + MaxWi% + 300
For i% = 1 To 2
    txtOptionenAutomatik(i%).Left = txtOptionenAutomatik(i% - 1).Left
Next i%

lblOptionenAutomatikMinuten(0).Left = txtOptionenAutomatik(0).Left + txtOptionenAutomatik(0).Width + 150
lblOptionenAutomatikMinuten(0).Top = txtOptionenAutomatik(0).Top
For i% = 1 To 2
    lblOptionenAutomatikMinuten(i%).Left = lblOptionenAutomatikMinuten(i% - 1).Left
    lblOptionenAutomatikMinuten(i%).Top = txtOptionenAutomatik(i%).Top
Next i%



fmeOptionenAutomatikBestVors.Left = lblOptionenAutomatik(0).Left
fmeOptionenAutomatikBestVors.Top = lblOptionenAutomatik(2).Top + lblOptionenAutomatik(2).Height + 450

txtOptionenAutomatikBestVors(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    txtOptionenAutomatikBestVors(i%).Top = txtOptionenAutomatikBestVors(i% - 1).Top + txtOptionenAutomatikBestVors(i% - 1).Height + 90
Next i%

chkOptionenAutomatikBestVors(0).Left = wpara.LinksX
chkOptionenAutomatikBestVors(0).Top = txtOptionenAutomatikBestVors(0).Top
For i% = 1 To 1
    chkOptionenAutomatikBestVors(i%).Left = chkOptionenAutomatikBestVors(i% - 1).Left
    chkOptionenAutomatikBestVors(i%).Top = txtOptionenAutomatikBestVors(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = chkOptionenAutomatikBestVors(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtOptionenAutomatikBestVors(0).Left = txtOptionenAutomatik(0).Left - fmeOptionenAutomatikBestVors.Left
For i% = 1 To 1
    txtOptionenAutomatikBestVors(i%).Left = txtOptionenAutomatikBestVors(i% - 1).Left
Next i%

lblOptionenAutomatikBestVorsMinuten(0).Left = txtOptionenAutomatikBestVors(0).Left + txtOptionenAutomatikBestVors(0).Width + 150
lblOptionenAutomatikBestVorsMinuten(0).Top = txtOptionenAutomatikBestVors(0).Top
For i% = 1 To 1
    lblOptionenAutomatikBestVorsMinuten(i%).Left = lblOptionenAutomatikBestVorsMinuten(i% - 1).Left
    lblOptionenAutomatikBestVorsMinuten(i%).Top = txtOptionenAutomatikBestVors(i%).Top
Next i%

fmeOptionenAutomatikBestVors.Width = lblOptionenAutomatikBestVorsMinuten(0).Left + lblOptionenAutomatikBestVorsMinuten(0).Width + 2 * wpara.LinksX
fmeOptionenAutomatikBestVors.Height = txtOptionenAutomatikBestVors(1).Top + txtOptionenAutomatikBestVors(1).Height + 2 * wpara.TitelY

Hoehe2% = fmeOptionenAutomatikBestVors.Top + fmeOptionenAutomatikBestVors.Height + 180
Breite2% = fmeOptionenAutomatikBestVors.Width + 2 * wpara.LinksX

'------------------------

tabOptionen.Tab = 6

For i% = 0 To 2
    lblOptionenDirektBezug(i%).Left = wpara.LinksX
Next i%
lblOptionenDirektBezug(0).Top = 2 * wpara.TitelY + 60
For i% = 1 To 2
    lblOptionenDirektBezug(i%).Top = lblOptionenDirektBezug(i% - 1).Top + lblOptionenDirektBezug(i% - 1).Height + 150
Next i%

xpos% = lblOptionenDirektBezug(0).Left + lblOptionenDirektBezug(0).Width
ydiff% = (txtOptionenDirektBezug(0).Height - lblOptionenDirektBezug(0).Height) / 2

For i% = 0 To 2
    txtOptionenDirektBezug(i%).Left = xpos%
    txtOptionenDirektBezug(i%).Top = lblOptionenDirektBezug(i%).Top - ydiff%
    lblOptionenDirektBezug2(i%).Top = lblOptionenDirektBezug(i%).Top
    lblOptionenDirektBezug2(i%).Left = txtOptionenDirektBezug(i%).Left + txtOptionenDirektBezug(i%).Width + 90
Next i%

fmeOptionenDirektBezug.Left = wpara.LinksX
fmeOptionenDirektBezug.Top = 900
fmeOptionenDirektBezug.Width = lblOptionenDirektBezug2(2).Left + lblOptionenDirektBezug2(2).Width + 2 * wpara.LinksX
fmeOptionenDirektBezug.Height = txtOptionenDirektBezug(2).Top + txtOptionenDirektBezug(2).Height + wpara.TitelY


For i% = 0 To 2
    chkOptionenDirektBezug2(i%).Left = wpara.LinksX
Next i%
chkOptionenDirektBezug2(0).Top = 2 * wpara.TitelY + 60
For i% = 1 To 2
    chkOptionenDirektBezug2(i%).Top = chkOptionenDirektBezug2(i% - 1).Top + chkOptionenDirektBezug2(i% - 1).Height + 90
Next i%

fmeOptionenDirektBezug2.Left = wpara.LinksX
fmeOptionenDirektBezug2.Top = fmeOptionenDirektBezug.Top + fmeOptionenDirektBezug.Height + 150
fmeOptionenDirektBezug2.Width = fmeOptionenDirektBezug.Width
fmeOptionenDirektBezug2.Height = chkOptionenDirektBezug2(2).Top + chkOptionenDirektBezug2(2).Height + wpara.TitelY

Hoehe3% = fmeOptionenDirektBezug2.Top + fmeOptionenDirektBezug2.Height + 180
If (Hoehe3% > Hoehe1%) Then
    Hoehe1% = Hoehe3%
End If
'------------------------



If (Hoehe1% > Hoehe2%) Then
    tabOptionen.Height = Hoehe1%
Else
    tabOptionen.Height = Hoehe2%
End If
If (Breite1% > Breite2%) Then
    tabOptionen.Width = Breite1%
Else
    tabOptionen.Width = Breite2%
End If
tabOptionen.Width = tabOptionen.Width + 5 * wpara.LinksX


cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = tabOptionen.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
If (Breite1% < 0) Then Breite1% = 0
Me.Left = Breite1%
Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
If (Hoehe1% < 0) Then Hoehe1% = 0
Me.Top = Hoehe1%

tabOptionen.Tab = 0
Call TabDisable
flxOptionen(0).Visible = True
cmdF2.Visible = True
cmdF5.Visible = True

Call DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (ActiveControl.Name = txtSchwellwerte(0).Name) Or (ActiveControl.Name = txtOptionenDirektBezug(0).Name) Then
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
        Beep
        KeyAscii = 0
    End If
End If
Call DefErrPop
End Sub

Private Sub tabOptionen_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tabOptionen_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabOptionen.Tab)

'Select Case tabOptionen.Tab
'    Case 0
'        flxOptionen(0).SetFocus
'    Case 1
'        flxOptionen(1).SetFocus
'    Case 2
'        flxOptionen(2).SetFocus
'    Case 3
'        txtOptionenAutomatik(0).SetFocus
'End Select

Call DefErrPop
End Sub

Sub OptionenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OptionenBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, l%, ind%, ind2%
Dim h$, h2$, h3$, lief12$(1), s$

lstOptionen.Visible = False
lstOptionenMulti.Visible = False
txtOptionen.Visible = False

For j% = 0 To 2
    With flxOptionen(j%)
        .Rows = 1
        If (j% = 0) Then
            For i% = 0 To (AnzKontrollen% - 1)
                h$ = RTrim$(Kontrollen(i%).bedingung.wert1)
                h$ = h$ + vbTab + RTrim$(Kontrollen(i%).bedingung.op)
                h$ = h$ + vbTab + RTrim$(Kontrollen(i%).bedingung.wert2)
                If (Kontrollen(i%).Send = "J") Then
                    h$ = h$ + vbTab + "Senden"
                ElseIf (Kontrollen(i%).Send = "N") Then
                    h$ = h$ + vbTab + "Nicht Senden"
                Else
                    h$ = h$ + vbTab + "und"
                End If
                .AddItem h$
            Next i%
            .Rows = MAX_KONTROLLEN
        ElseIf (j% = 1) Then
            For i% = 0 To (AnzZuordnungen% - 1)
                h$ = RTrim$(Zuordnungen(i%).bedingung.wert1)
                h$ = h$ + vbTab + RTrim$(Zuordnungen(i%).bedingung.op)
                h$ = h$ + vbTab + RTrim$(Zuordnungen(i%).bedingung.wert2)
    
                h$ = h$ + vbTab
                h2$ = ""
                For k% = 0 To 19
                    If (Zuordnungen(i%).lief(k%) > 0) Then
                        h2$ = h2$ + Mid$(Str$(Zuordnungen(i%).lief(k%)), 2) + ","
                    Else
                        Exit For
                    End If
                Next k%
                If (k% = 1) Then
                    If (Zuordnungen(i%).lief(0) = 999) Then
                        h3$ = "und"
                    ElseIf (Zuordnungen(i%).lief(0) = 255) Then
                        h3$ = "Naechstliefernder"
                    Else
                        lif.GetRecord (Zuordnungen(i%).lief(0) + 1)
                        h3$ = RTrim$(lif.kurz)
                    End If
                    h$ = h$ + h3$
                ElseIf (k% > 1) Then
                    h$ = h$ + "mehrere (" + Mid$(Str$(k%), 2) + ")"
                End If
                h$ = h$ + vbTab + h2$
                .AddItem h$
            Next i%
            .Rows = MAX_ZUORDNUNGEN
        ElseIf (j% = 2) Then
            For i% = 0 To (AnzRufzeiten% - 1)
                lif.GetRecord (Rufzeiten(i%).Lieferant + 1)
                h$ = RTrim$(lif.kurz)
                h$ = h$ + " (" + Mid$(Str$(Rufzeiten(i%).Lieferant), 2) + ")"
    
                h$ = h$ + vbTab
                If (Rufzeiten(i%).RufZeit < 9000) Then
                    h$ = h$ + Format$(Rufzeiten(i%).RufZeit \ 100, "00")
                    h$ = h$ + ":" + Format$(Rufzeiten(i%).RufZeit Mod 100, "00")
                End If
                h$ = h$ + vbTab
                If (Rufzeiten(i%).LieferZeit < 9000) Then
                    h$ = h$ + Format$(Rufzeiten(i%).LieferZeit \ 100, "00")
                    h$ = h$ + ":" + Format$(Rufzeiten(i%).LieferZeit Mod 100, "00")
                End If
    
                h$ = h$ + vbTab
                h2$ = ""
                For k% = 0 To 6
                    ind% = Rufzeiten(i%).WoTag(k%)
                    If (ind% > 0) Then
                        h2$ = h2$ + Mid$(Str$(ind%), 2) + ","
                    Else
                        Exit For
                    End If
                Next k%
                If (k% = 0) Then
                ElseIf (k% = 1) Then
                    ind% = Rufzeiten(i%).WoTag(0)
                    h$ = h$ + WochenTag$(ind% - 1)
                ElseIf (k% < 4) Then
                    For l% = 0 To (k% - 1)
                        ind% = Rufzeiten(i%).WoTag(l%)
                        h$ = h$ + WochenTag$(ind% - 1) + ","
                    Next l%
                Else
                    For l% = 0 To (k% - 1)
                        ind% = Rufzeiten(i%).WoTag(l%)
                        h$ = h$ + Left$(WochenTag$(ind% - 1), 2) + ","
                    Next l%
                End If
                h$ = h$ + vbTab + h2$
                h$ = h$ + vbTab + Rufzeiten(i%).AuftragsErg
                h$ = h$ + vbTab + Rufzeiten(i%).AuftragsArt
                If (Rufzeiten(i%).Aktiv = "J") Then
                    h$ = h$ + vbTab + "ja"
                Else
                    h$ = h$ + vbTab + "nein"
                End If
                .AddItem h$
            Next i%
            .Rows = MAX_RUFZEITEN
        Else
            For i% = 0 To (AnzFeiertage% - 1)
                h$ = " "
                If (Feiertage(i%).Aktiv = "J") Then h$ = Chr$(214)
                h$ = h$ + vbTab + RTrim$(Feiertage(i%).Name) + vbTab + RTrim$(Feiertage(i%).KalenderTag)
                .AddItem h$
            Next i%
'            .Rows = MAX_FEIERTAGE
            
            .FillStyle = flexFillRepeat
            .row = 1
            .col = 0
            .RowSel = .Rows - 1
            .ColSel = .col
            .CellFontName = "Symbol"
            .FillStyle = flexFillSingle
        End If
        
        .row = 1
        .col = 0
    End With
Next j%

With flxFeiertage
    .Rows = 1
    For i% = 0 To (AnzFeiertage% - 1)
        h$ = " "
        If (Feiertage(i%).Aktiv = "J") Then h$ = Chr$(214)
        h$ = h$ + vbTab + RTrim$(Feiertage(i%).Name) + vbTab + RTrim$(Feiertage(i%).KalenderTag)
        .AddItem h$
    Next i%
'   .Rows = MAX_FEIERTAGE
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 0
End With

With flxAllgLieferanten
    .Rows = 1
    h$ = GhRufzeiten$
    Do
        ind% = InStr(h$, ",")
        If (ind% <= 0) Then
            Exit Do
        Else
            h2$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        
            ind2% = 0
            Do
                ind% = InStr(ind2% + 1, h2$, "-")
                If (ind% > 0) Then
                    ind2% = ind%
                Else
                    ind% = ind2%
                    Exit Do
                End If
            Loop
            
            h3$ = " "
            If (InStr(GhAllgAngebote$, Mid$(h2$, ind% + 1)) > 0) Then h3$ = Chr$(214)
            h3$ = h3$ + vbTab + Left$(h2$, ind% - 1) + vbTab + Mid$(h2$, ind% + 1)
            .AddItem h3$
        End If
    Loop
    
    If (.Rows = 1) Then .AddItem " "
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 0
End With

With flxAbsagenKz
    .Rows = 1
    .AddItem vbTab + "0R"
    .AddItem vbTab + "00"
    .AddItem vbTab + "BS"
    .AddItem vbTab + "AP"
    .AddItem vbTab + "AV"
    .AddItem vbTab + "KF"
    .AddItem vbTab + "NE"
    .AddItem vbTab + "NG"
    .AddItem vbTab + "NL"
    .AddItem vbTab + "OT"
    .AddItem vbTab + "KS"

    If (.Rows = 1) Then .AddItem " "
    
    If (Right$(AbsagenMitNL$, 1) <> ",") Then
        AbsagenMitNL$ = AbsagenMitNL$ + ","
    End If
    
    For i% = 1 To (.Rows - 1)
        h3$ = " "
        If (InStr(AbsagenMitNL$, .TextMatrix(i%, 1) + ",") > 0) Then h3$ = Chr$(214)
        .TextMatrix(i%, 0) = h3$
    Next i%
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 0
End With

txtAutomatenWare.text = AutomatenLac$
With flxAutomatenWare
    .Rows = 1
    If (AutomatenLiefs$ <> "") Then
        h$ = Trim(AutomatenLiefs$)
        If (Left$(h$, 1) = ",") Then
            h$ = Mid$(h$, 2)
        End If
        
        Do
            ind% = InStr(h$, ",")
            If (ind% > 0) Then
                h2$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                ind% = InStr(h2$, "-")
                If (ind% > 0) Then
                    lief12$(0) = Format(Val(Left$(h2$, ind% - 1)), "0")
                    lief12$(1) = Format(Val(Mid$(h2$, ind% + 1)), "0")
                    
                    h3$ = ""
                    For j% = 0 To 1
                        s$ = "(" + lief12$(j%) + ")"
                        For i% = 1 To AnzLiefNamen%
                            h2$ = LiefNamen$(i% - 1)
                            If (InStr(h2$, s$) > 0) Then
                                h3$ = h3$ + h2$
                                Exit For
                            End If
                        Next i%
                        h3$ = h3$ + vbTab
                    Next j%
                    
                    .AddItem h3$
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    If (.Rows < 11) Then
        .Rows = 11
    End If
End With

txtOptionenAutomatik(0).text = Str$(AnzMinutenWarnung%)
txtOptionenAutomatik(1).text = Str$(AnzMinutenWarten%)
txtOptionenAutomatik(2).text = Str$(AnzMinutenVerspaetung%)

chkOptionenAutomatikBestVors(0).Value = Abs(BestVorsKomplett%)
chkOptionenAutomatikBestVors(1).Value = Abs(BestVorsPeriodisch%)
txtOptionenAutomatikBestVors(0).text = Str$(BestVorsKomplettMinuten%)
txtOptionenAutomatikBestVors(0).Visible = Abs(BestVorsKomplett%)
txtOptionenAutomatikBestVors(1).text = Str$(BestVorsPeriodischMinuten%)
txtOptionenAutomatikBestVors(1).Visible = Abs(BestVorsPeriodisch%)
lblOptionenAutomatikBestVorsMinuten(0).Visible = Abs(BestVorsKomplett%)
lblOptionenAutomatikBestVorsMinuten(1).Visible = Abs(BestVorsPeriodisch%)

chkSchwellwerteAktiv.Value = Abs(SchwellwertAktiv%)
txtSchwellwerte(0).text = Format(SchwellwertMinuten%, "0")
txtSchwellwerte(1).text = Format(SchwellwertSicherheit%, "0")
txtSchwellwerte(2).text = Format(SchwellwertWarnungProz%, "0")
txtSchwellwerte(3).text = Format(SchwellwertWarnungAb%, "0")
chkSchwellwerte(0).Value = Abs(SchwellwertVorab%)
chkSchwellwerte(1).Value = Abs(SchwellwertGlaetten%)

txtOptionenDirektBezug(0).text = Format(DirektBezugSendMinuten%, "0")
txtOptionenDirektBezug(1).text = Format(DirektBezugKontrollenMinunten%, "0")
txtOptionenDirektBezug(2).text = Format(DirektBezugWarnungMinuten%, "0")

chkOptionenDirektBezug2(0).Value = Abs(MitLagerstandCalc%)
chkOptionenDirektBezug2(1).Value = Abs(WuPruefung%)
chkOptionenDirektBezug2(2).Value = Abs(Bm0Anzeigen%)

Call DefErrPop
End Sub

Private Sub chkOptionenAutomatikBestVors_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkOptionenAutomatikBestVors_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    BestVorsKomplett% = chkOptionenAutomatikBestVors(0).Value
Else
    BestVorsPeriodisch% = chkOptionenAutomatikBestVors(1).Value
End If
txtOptionenAutomatikBestVors(0).Visible = Abs(BestVorsKomplett%)
txtOptionenAutomatikBestVors(1).Visible = Abs(BestVorsPeriodisch%)
lblOptionenAutomatikBestVorsMinuten(0).Visible = Abs(BestVorsKomplett%)
lblOptionenAutomatikBestVorsMinuten(1).Visible = Abs(BestVorsPeriodisch%)

Call DefErrPop
End Sub

Private Sub lstOptionen_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lstOptionen_DblClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call cmdOk_Click
Call DefErrPop
End Sub

Private Sub flxOptionen_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen_DblClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call cmdOk_Click
Call DefErrPop
End Sub

Private Sub txtOptionenAutomatik_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionenAutomatik_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtOptionenAutomatik(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtOptionenAutomatikbestvors_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionenAutomatikbestvors_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (tabOptionen.Tab <> 3) Then
'    flxOptionen(tabOptionen.Tab).SetFocus
'End If

With txtOptionenAutomatikBestVors(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub chkSchwellwerteAktiv_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkSchwellwerteAktiv_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

fmeSchwellwerte.Visible = chkSchwellwerteAktiv.Value

Call DefErrPop
End Sub

Private Sub txtSchwellwerte_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtSchwellwerte_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With txtSchwellwerte(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtOptionenDirektBezug_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionenDirektBezug_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtOptionenDirektBezug(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub flxFeiertage_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxFeiertage_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
With flxFeiertage
    .HighLight = flexHighlightAlways
    .col = 0
    .ColSel = .Cols - 1
End With
Call DefErrPop
End Sub

Private Sub flxFeiertage_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxFeiertage_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
flxFeiertage.HighLight = flexHighlightNever
Call DefErrPop
End Sub

Private Sub flxAllgLieferanten_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAllgLieferanten_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
With flxAllgLieferanten
    .HighLight = flexHighlightAlways
    .col = 0
    .ColSel = .Cols - 1
End With
Call DefErrPop
End Sub

Private Sub flxAllgLieferanten_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAllgLieferanten_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
flxAllgLieferanten.HighLight = flexHighlightNever
Call DefErrPop
End Sub

Private Sub flxAbsagenKz_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAbsagenKz_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
With flxAbsagenKz
    .HighLight = flexHighlightAlways
    .col = 0
    .ColSel = .Cols - 1
End With
Call DefErrPop
End Sub

Private Sub flxAbsagenKz_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAbsagenKz_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
flxAbsagenKz.HighLight = flexHighlightNever
Call DefErrPop
End Sub

Private Sub txtAutomatenWare_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAutomatenWare_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With txtAutomatenWare
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub AuslesenFlexKontrollen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexKontrollen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, ind%
Dim l&
Dim h$, Send$, BetrLief$, lief1$, Lief2$, BetrTage$, tag$, Aktiv$

With flxOptionen(0)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Kontrollen(j%).bedingung.wert1 = h$
            Kontrollen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            Kontrollen(j%).bedingung.wert2 = h$
            
            Send$ = RTrim$(.TextMatrix(i%, 3))
            If (Send$ = "Nicht Senden") Then
                Kontrollen(j%).Send = "N"
            ElseIf (Send$ = "Senden") Then
                Kontrollen(j%).Send = "J"
            Else
                Kontrollen(j%).Send = "U"
            End If
            j% = j% + 1
        End If
    Next i%
End With
AnzKontrollen% = j%

With flxOptionen(1)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Zuordnungen(j%).bedingung.wert1 = h$
            Zuordnungen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            Zuordnungen(j%).bedingung.wert2 = RTrim$(.TextMatrix(i%, 2))
            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 4)))
            
            If (BetrLief$ = "999") Then
                Zuordnungen(j%).lief(0) = 999
                k% = 1
            Else
                For k% = 0 To 19
                    If (BetrLief$ = "") Then Exit For
                    
                    ind% = InStr(BetrLief$, ",")
                    If (ind% > 0) Then
                        Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                        BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                    Else
                        Lief2$ = BetrLief$
                        BetrLief$ = ""
                    End If
                    If (Lief2$ <> "") Then
                        Zuordnungen(j%).lief(k%) = Val(Lief2$)
                    End If
                Next k%
            End If
            Do
                If (k% > 19) Then Exit Do
                Zuordnungen(j%).lief(k%) = 0
                k% = k% + 1
            Loop
            j% = j% + 1
        End If
    Next i%
End With
AnzZuordnungen% = j%

With flxOptionen(2)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = LTrim$(RTrim$(.TextMatrix(i%, 0)))
        If (h$ <> "") Then
            ind% = InStr(h$, "(")
            If (ind% > 0) Then
                h$ = Mid$(h$, ind% + 1)
                ind% = InStr(h$, ")")
                h$ = Left$(h$, ind% - 1)
                Rufzeiten(j%).Lieferant = Val(h$)
            End If
            
            h$ = RTrim$(.TextMatrix(i%, 1))
            If (h$ <> "") Then
                h$ = Left$(h$, 2) + Mid$(h$, 4)
            End If
            Rufzeiten(j%).RufZeit = Val(h$)
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            If (h$ <> "") Then
                h$ = Left$(h$, 2) + Mid$(h$, 4)
            End If
            Rufzeiten(j%).LieferZeit = Val(h$)
            
            BetrTage$ = LTrim$(RTrim$(.TextMatrix(i%, 4)))
            For k% = 0 To 6
                If (BetrTage$ = "") Then Exit For
                
                ind% = InStr(BetrTage$, ",")
                If (ind% > 0) Then
                    tag$ = RTrim$(Left$(BetrTage$, ind% - 1))
                    BetrTage$ = LTrim$(Mid$(BetrTage$, ind% + 1))
                Else
                    tag$ = BetrTage$
                    BetrTage$ = ""
                End If
                If (tag$ <> "") Then
                    Rufzeiten(j%).WoTag(k%) = Val(tag$)
                End If
            Next k%
            Do
                If (k% > 6) Then Exit Do
                Rufzeiten(j%).WoTag(k%) = 0
                k% = k% + 1
            Loop
            
            h$ = Left$(LTrim$(RTrim$(.TextMatrix(i%, 5))) + Space$(2), 2)
            Rufzeiten(j%).AuftragsErg = h$
            
            h$ = Left$(LTrim$(RTrim$(.TextMatrix(i%, 6))) + Space$(2), 2)
            Rufzeiten(j%).AuftragsArt = h$
            
            Aktiv$ = RTrim$(.TextMatrix(i%, 7))
            If (Aktiv$ = "ja") Then
                Rufzeiten(j%).Aktiv = "J"
            Else
                Rufzeiten(j%).Aktiv = "N"
            End If
            
            j% = j% + 1
        End If
    Next i%
End With
AnzRufzeiten% = j%

With flxFeiertage
    For i% = 1 To (.Rows - 1)
        h$ = Trim$(.TextMatrix(i%, 0))
        Aktiv$ = "N"
        If (h$ <> "") Then Aktiv$ = "J"
        Feiertage(i% - 1).Aktiv = Aktiv$
    Next i%
End With

With flxAllgLieferanten
    GhAllgAngebote$ = ""
    For i% = 1 To (.Rows - 1)
        h$ = Trim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            h$ = Trim$(.TextMatrix(i%, 2))
            If (h$ <> "") Then GhAllgAngebote$ = GhAllgAngebote$ + h$ + ","
        End If
    Next i%
    l& = WritePrivateProfileString("Bestellung", "AllgemeineAngebote", GhAllgAngebote$, INI_DATEI)
End With

With flxAbsagenKz
    AbsagenMitNL$ = ""
    For i% = 1 To (.Rows - 1)
        h$ = Trim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            h$ = Trim$(.TextMatrix(i%, 1))
            If (h$ <> "") Then AbsagenMitNL$ = AbsagenMitNL$ + h$ + ","
        End If
    Next i%
    l& = WritePrivateProfileString("Bestellung", "AbsagenMitNL", AbsagenMitNL$, INI_DATEI)
End With

AutomatenLiefs$ = ""
AutomatenLac$ = Trim(txtAutomatenWare.text)
If (AutomatenLac$ <> "") Then
    With flxAutomatenWare
        h$ = ","
        For i% = 1 To (.Rows - 1)
            lief1$ = Trim$(.TextMatrix(i%, 0))
            Lief2$ = Trim$(.TextMatrix(i%, 1))
            If (lief1$ <> "") And (Lief2$ <> "") Then
                ind% = InStr(lief1$, "(")
                If (ind% > 0) Then
                    lief1$ = Mid$(lief1$, ind% + 1)
                    ind% = InStr(lief1$, ")")
                    lief1$ = Left$(lief1$, ind% - 1)
                End If
            
                ind% = InStr(Lief2$, "(")
                If (ind% > 0) Then
                    Lief2$ = Mid$(Lief2$, ind% + 1)
                    ind% = InStr(Lief2$, ")")
                    Lief2$ = Left$(Lief2$, ind% - 1)
                End If
                
                h$ = h$ + Format(Val(lief1$), "000") + "-" + Format(Lief2$, "000") + ","
            End If
        Next i%
    End With
    AutomatenLiefs$ = h$
End If
h$ = AutomatenLac$ + AutomatenLiefs$
l& = WritePrivateProfileString("Bestellung", "AutomatenLieferanten", h$, INI_DATEI)


AnzMinutenWarnung% = Val(txtOptionenAutomatik(0).text)
AnzMinutenWarten% = Val(txtOptionenAutomatik(1).text)
AnzMinutenVerspaetung% = Val(txtOptionenAutomatik(2).text)
BestVorsKomplett% = chkOptionenAutomatikBestVors(0).Value
BestVorsKomplettMinuten% = Val(txtOptionenAutomatikBestVors(0).text)
BestVorsPeriodisch% = chkOptionenAutomatikBestVors(1).Value
BestVorsPeriodischMinuten% = Val(txtOptionenAutomatikBestVors(1).text)

SchwellwertAktiv% = chkSchwellwerteAktiv.Value
SchwellwertMinuten% = Val(txtSchwellwerte(0).text)
SchwellwertSicherheit% = Val(txtSchwellwerte(1).text)
SchwellwertWarnungProz% = Val(txtSchwellwerte(2).text)
SchwellwertWarnungAb% = Val(txtSchwellwerte(3).text)
SchwellwertVorab% = chkSchwellwerte(0).Value
SchwellwertGlaetten% = chkSchwellwerte(1).Value

DirektBezugSendMinuten% = Val(txtOptionenDirektBezug(0).text)
DirektBezugKontrollenMinunten% = Val(txtOptionenDirektBezug(1).text)
DirektBezugWarnungMinuten% = Val(txtOptionenDirektBezug(2).text)

MitLagerstandCalc% = chkOptionenDirektBezug2(0).Value
WuPruefung% = chkOptionenDirektBezug2(1).Value
Bm0Anzeigen% = chkOptionenDirektBezug2(2).Value

Call DefErrPop
End Sub

Sub AnzeigeContainer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AnzeigeContainer")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim h$, h2$
Dim c As Control

On Error Resume Next
If (prot% = True) Then
    For Each c In Controls
        h$ = c.Name
        h2$ = ""
        h2$ = RTrim(Format(c.Index, "00"))
        If (h2$ <> "") Then
            h$ = h$ + "(" + h2$ + ")"
        End If
        h$ = Left$(h$ + Space$(40), 40)
        h$ = h$ + c.Container.Name
        h2$ = ""
        h2$ = RTrim(Format(c.Container.Index, "00"))
        If (h2$ <> "") Then
            h$ = h$ + "(" + h2$ + ")"
        End If
        Print #PROTOKOLL%, h$
    Next c
End If

Call DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabDisable")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

For i% = 0 To 2
    flxOptionen(i%).Visible = False
Next i%
flxFeiertage.Visible = False
flxAllgLieferanten.Visible = False
flxAbsagenKz.Visible = False

cmdF2.Visible = False
cmdF5.Visible = False
For i% = 0 To 1
    cmdChange(i%).Visible = False
Next i%

For i% = 0 To 2
    lblOptionenAutomatik(i%).Visible = False
    txtOptionenAutomatik(i%).Visible = False
    lblOptionenAutomatikMinuten(i%).Visible = False
Next i%
fmeOptionenAutomatikBestVors.Visible = False

chkSchwellwerteAktiv.Visible = False
fmeSchwellwerte.Visible = False

fmeOptionenDirektBezug.Visible = False
fmeOptionenDirektBezug2.Visible = False

lblAutomatenWare.Visible = False
txtAutomatenWare.Visible = False
flxAutomatenWare.Visible = False

Call DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabEnable")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

If (hTab% = 3) Then
    For i% = 0 To 2
        lblOptionenAutomatik(i%).Visible = True
        txtOptionenAutomatik(i%).Visible = True
        lblOptionenAutomatikMinuten(i%).Visible = True
    Next i%
    fmeOptionenAutomatikBestVors.Visible = True
    txtOptionenAutomatik(0).SetFocus
ElseIf (hTab% = 4) Then
    chkSchwellwerteAktiv.Visible = True
    fmeSchwellwerte.Visible = chkSchwellwerteAktiv.Value
'    fmeSchwellwerte.Visible = True
    chkSchwellwerteAktiv.SetFocus
ElseIf (hTab% = 5) Then
    flxFeiertage.Visible = True
    flxFeiertage.col = 0
    flxFeiertage.ColSel = flxFeiertage.Cols - 1
    flxFeiertage.SetFocus
    flxAllgLieferanten.Visible = True
    If (para.Land = "A") Then
        flxAbsagenKz.Visible = True
    End If
ElseIf (hTab% = 6) Then
    fmeOptionenDirektBezug.Visible = True
    fmeOptionenDirektBezug2.Visible = True
    txtOptionenDirektBezug(0).SetFocus
ElseIf (hTab% = 7) Then
    lblAutomatenWare.Visible = True
    txtAutomatenWare.Visible = True
    flxAutomatenWare.Visible = True
    txtAutomatenWare.SetFocus
Else
    flxOptionen(hTab%).Visible = True
    flxOptionen(hTab%).SetFocus

    cmdF2.Visible = True
    cmdF5.Visible = True
    For i% = 0 To 1
        cmdChange(i%).Visible = True
    Next i%
End If

Call DefErrPop
End Sub

Sub EditOptionenLst()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenLst")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, hTab%, row%, col%, xBreit%, ind%, aRow%
Dim s$, h$

hTab% = tabOptionen.Tab
If (hTab% = 4) Then hTab% = 3
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))

xBreit% = False
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            

With frmEdit.lstEdit
    .Clear
    Select Case col%
        Case 0
            If (hTab% < 2) Then
                .AddItem "Absagen"
                .AddItem "Dauer-Absagen"
                .AddItem "Anfragen"
                .AddItem "akzept.Angebote"
                .AddItem "Angebotshinweise"
                .AddItem "Besorger"
                .AddItem "Text-Besorger"
                .AddItem "Bewertung"
                .AddItem "BTM"
                .AddItem "Kühl/Kalt"
                .AddItem "Manuelle"
                .AddItem "Ladenhüter"
                .AddItem "Lagerartikel"
                .AddItem "def.Lagerartikel"
                .AddItem "Lagerartikel neg.LS"
                .AddItem "Schnelldreher"
                .AddItem "def.Schnelldreher"
                .AddItem "Originale"
                .AddItem "Importe"
'                .AddItem "Doppeltkontrolle"
                .AddItem "Selbstangelegte"
                .AddItem "Interne Streichung"
                .AddItem "Außer Handel"
                .AddItem "preisg.Artikel"
                .AddItem "preisg.Art. vorh."
                .AddItem "Bestellung Partner"
                .AddItem "Interim"
                .AddItem "AM RX"
                .AddItem "AM non RX"
                .AddItem "NichtArzneimittel"
                .AddItem String$(50, "-")
                .AddItem "BM"
                .AddItem "EK"
                .AddItem "VK"
                .AddItem "Zeilenwert"
                .AddItem "übl.Lieferant"
                .AddItem "Lagerstand"
                .AddItem String$(50, "-")
                .AddItem "Hersteller"
                .AddItem "Lieferant"
                .AddItem "Lagercode"
                .AddItem "Lagercode 2st."
                .AddItem "Warengruppe"
                .AddItem String$(50, "-")
                .AddItem "Uhrzeit"
    '                            .AddItem "Zwischenbestellung"
            ElseIf (hTab% = 2) Then
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
            End If
            
        Case 1
            If (hTab% = 3) Then
                .AddItem "geöffnet"
                .AddItem "geschlossen"
            Else
                Select Case ZeilenTyp%(s$)
                    Case 0
                        .AddItem "vor"
                        .AddItem "nach"
                        If (s$ = "Besorger") Then
                            .AddItem "älter als"
                        End If
                    
                    Case 1
                        .AddItem "="
                        .AddItem "<>"
    
                    Case 3
                        .AddItem ">="
    
                    Case Else
                        .AddItem "<"
                        .AddItem "<="
                        .AddItem "="
                        .AddItem "<>"
                        .AddItem ">="
                        .AddItem ">"
        '                        .AddItem "?-fach"
                End Select
            End If
    
        Case 2
            If (s$ = "Lieferant") Then
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
            ElseIf (s$ = "Besorger") Then
                .AddItem "1 Tag"
                .AddItem "2 Tage"
                .AddItem "3 Tage"
                .AddItem "4 Tage"
                .AddItem "5 Tage"
                .AddItem "Eingabe"
            Else
                .AddItem "1"
                .AddItem "2"
                .AddItem "3"
                .AddItem "4"
                .AddItem "5"
                .AddItem "Eingabe"
                If (s$ = "BM") Then
                    .AddItem "BMopt"
                    .AddItem "?-fach"
                End If
            End If
    
        Case 3
            .AddItem "Nicht Senden"
            .AddItem "Senden"
            .AddItem String$(50, "-")
            .AddItem "und"
    
        Case 5
            If (hTab% = 2) Then
                If (para.Land = "A") Then
                    .AddItem "Normalfall (  )"
                    .AddItem "Kein Auftrag (KA)"
                    .AddItem "Rückruf erwünscht (RR)"
                    .AddItem "Später (SP)"
                    .AddItem "Kein Auftrag, Rückruf (KR)"
                Else
                    .AddItem "Zustellung Heute (ZH)"
                    .AddItem "Zustellung Morgen (ZM)"
                    .AddItem "Heute kein Auftrag (KA)"
                    .AddItem "Rückruf erbeten (RR)"
                End If
                xBreit% = True
            End If
        
        Case 6
            If (hTab% = 2) Then
                If (para.Land = "A") Then
                    .AddItem "Normalauftrag in Dekade (N0)"
                    .AddItem "Normalauftrag außer Dekade (NA)"
                    .AddItem "Testauftrag in Dekade (TE)"
                Else
                    .AddItem "Normalauftrag (  )"
                    .AddItem "Inventurliste (IN)"
                    .AddItem "Lochkarten (LK)"
                    .AddItem "Rückkauf-Anfrage (RK)"
                    .AddItem "SBL-Auftrag (SB)"
                    .AddItem "Sonder-Auftrag (SO)"
                    .AddItem "Stapel-Auftrag (ST)"
                    .AddItem "Test-Auftrag (TE)"
                    .AddItem "Verfalldatenliste (VD)"
                    .AddItem "Vorratskauf (VR)"
                    .AddItem "10er-Auftrag (ZE)"
                    .AddItem "Zeit-Auftrag (ZT)"
                    .AddItem "Zeilenwert-Auftrag (ZW)"
                End If
                xBreit% = True
            End If
        
        Case 7
            If (hTab% = 2) Then
                .AddItem "ja"
                .AddItem "nein"
            End If
    End Select


    .ListIndex = 0
    s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, col%))
    If (InStr(s$, "*") > 0) Then s$ = "?-fach"
    If (s$ <> "") Then
        If (col% = 5) Or (col% = 6) Then
            s$ = "(" + s$ + ")"
        End If
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h$ = .text
            If (col% = 4) Or (col% = 5) Then
                If (InStr(h$, s$) > 0) Then
                    Exit For
                End If
            ElseIf (s$ = h$) Then
                Exit For
            End If
        Next i%
    End If
    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(flxOptionen(hTab%).TopRow)  '1
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        If (xBreit%) Then
            .Width = TextWidth("Zeilenwert-Auftrag (ZW)") + wpara.FrmScrollHeight + 90
        Else
            .Width = flxOptionen(hTab%).ColWidth(col%)
        End If
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(flxOptionen(hTab%).TopRow)  '1
    End With
    With frmEdit.lstEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With
    
    frmEdit.Show 1
    
    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then

        h$ = EditTxt$
            With flxOptionen(hTab%)
                If (Left$(h$, 1) <> "-") Then
                    If (col% = 5) Or (col% = 6) Then
                        ind% = InStr(h$, "(")
                        If (ind% > 0) Then
                            h$ = Mid$(h$, ind% + 1, 2)
                        End If
                    End If
                    .TextMatrix(row%, col%) = h$
                    If (.col < .Cols - 2) Then .col = .col + 1
                    If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
                End If
            End With
'        End If
        
    End If

End With

Call DefErrPop
End Sub
                            
                            
Sub EditOptionenLstMulti()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenLstMulti")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, l%, hTab%, row%, col%, ind%, aRow%
Dim s$, h$, BetrLief$, Lief2$

hTab% = tabOptionen.Tab
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.lstMultiEdit
    .Clear
    Select Case col%
        Case 3
            If (hTab% = 2) Then
                .AddItem "(keiner)"
                .AddItem "Montag"
                .AddItem "Dienstag"
                .AddItem "Mittwoch"
                .AddItem "Donnerstag"
                .AddItem "Freitag"
                .AddItem "Samstag"
                .AddItem "Sonntag"
                
            ElseIf (hTab% > 0) Then
                .AddItem "(keiner)"
                .AddItem String$(50, "-")
                .AddItem "Nächstliefernder (255)"
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
                .AddItem String$(50, "-")
                .AddItem "und"
            End If
    End Select

    For i% = 0 To (.ListCount - 1)
        .Selected(i%) = False
    Next i%

    
    Load frmEdit
    
     .ListIndex = 0
     
     If (hTab% < 2) Then
         BetrLief$ = LTrim$(RTrim$(flxOptionen(hTab%).TextMatrix(row%, 4)))
     Else
         BetrLief$ = LTrim$(RTrim$(flxOptionen(hTab%).TextMatrix(row%, 4)))
     End If
     
     If (BetrLief$ = "999") Or (BetrLief$ = "999,") Then
        j% = .ListCount - 1
        .ListIndex = j%
        .Selected(j%) = True
     Else
        For i% = 0 To 19
            If (BetrLief$ = "") Then Exit For
            
            ind% = InStr(BetrLief$, ",")
            If (ind% > 0) Then
                Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
            Else
                Lief2$ = BetrLief$
                BetrLief$ = ""
            End If
            
            If (Lief2$ <> "") Then
                ind% = Val(Lief2$)
                If (hTab% < 2) Then
                    h$ = "(" + Mid$(Str$(ind%), 2) + ")"
                    For j% = 0 To (.ListCount - 1)
                        .ListIndex = j%
                        s$ = .text
                        If (InStr(s$, h$) > 0) Then
                            .Selected(j%) = True
                            Exit For
                        End If
                    Next j%
                Else
                    .Selected(ind%) = True
                End If
            End If
        Next i%
    End If
     
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(flxOptionen(hTab%).TopRow)  '1
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen(hTab%).ColWidth(col%)
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(flxOptionen(hTab%).TopRow)  '1
    End With
    With frmEdit.lstMultiEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With

    frmEdit.Show 1

    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then
            
        flxOptionen(hTab%).TextMatrix(row%, col% + 1) = EditTxt$
        
        h$ = ""
        If (hTab% = 1) Then
            If (EditTxt$ = "und") Then
                flxOptionen(hTab%).TextMatrix(row%, col% + 1) = "999"
                h$ = "und"
            ElseIf (EditGef%(0) = 255) Then
                flxOptionen(hTab%).TextMatrix(row%, col% + 1) = "255"
                h$ = "Naechstliefernder"
            Else
                If (EditAnzGefunden% = 0) Then
                    h$ = ""
                ElseIf (EditAnzGefunden% = 1) Then
                    ind% = EditGef%(0)
                    lif.GetRecord (ind% + 1)
                    h$ = lif.kurz
                Else
                    h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
                End If
            End If
        Else
            If (EditAnzGefunden% = 0) Then
            ElseIf (EditAnzGefunden% = 1) Then
                ind% = EditGef%(0)
                h$ = h$ + WochenTag$(ind% - 1)
            ElseIf (EditAnzGefunden% < 4) Then
                For l% = 0 To (EditAnzGefunden% - 1)
                    ind% = EditGef%(l%)
                    h$ = h$ + WochenTag$(ind% - 1) + ","
                Next l%
            Else
                For l% = 0 To (EditAnzGefunden% - 1)
                    ind% = EditGef%(l%)
                    h$ = h$ + Left$(WochenTag$(ind% - 1), 2) + ","
                Next l%
            End If
        End If

        With flxOptionen(hTab%)
            .TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
        
    End If

End With

Call DefErrPop
End Sub
                            
Sub EditOptionenTxt()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenTxt")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, hTab%, row%, col%, aRow%, uhr%, st%, min%, iZeilenTyp%
Dim s$, h$

EditModus% = 2

hTab% = tabOptionen.Tab
If (hTab% = 4) Then hTab% = 3
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
iZeilenTyp% = ZeilenTyp%(s$)
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.txtEdit
    Select Case col%
        Case 0
            If (hTab% = 3) Then
                .MaxLength = 4
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .text = h$
                EditModus% = 9
            End If
        Case 1
            If (hTab% = 2) Then
                .MaxLength = 4
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
            End If
        Case 2
            If (hTab% = 2) Then
                .MaxLength = 4
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
            Else
                If (iZeilenTyp% = 0) Then
                    .MaxLength = 4
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
                ElseIf (iZeilenTyp% = 1) Then
                    .MaxLength = 5
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .text = Right$("     " + h$, 5)
                    EditModus% = 1
                Else
                    .MaxLength = 4
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .text = Right$("    " + h$, 4)
                    EditModus% = 0
                End If
            End If
    End Select


    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen(hTab%).ColWidth(col%)
        .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
    End With
    With frmEdit.txtEdit
        .Width = frmEdit.ScaleWidth
'            .Height = frmEdit.ScaleHeight
        .Left = 0
        .Top = 0
        .BackColor = vbWhite
        .Visible = True
    End With
    
    
    frmEdit.Show 1
    
    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then

        If (col% = 0) Then
            h$ = Trim(EditTxt$)
        ElseIf (col% = 1) Or ((col% = 2) And ((hTab% = 2) Or (iZeilenTyp% = 0))) Then
            uhr% = Val(EditTxt$)
            st% = uhr% \ 100
            min% = uhr% Mod 100
            h$ = Format(st%, "00") + ":" + Format(min%, "00")
        ElseIf (iZeilenTyp% = 1) Then
            h$ = UCase$(Trim$(EditTxt$))
        Else
            h$ = Mid$(Str(Val(EditTxt$)), 2)
        End If
        With flxOptionen(hTab%)
            flxOptionen(hTab%).TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
    End If

End With

Call DefErrPop
End Sub

Sub EditAutomatenLiefs()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditAutomatenLiefs")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, hTab%, row%, col%, xBreit%, ind%, aRow%
Dim s$, h$

xBreit% = False

With flxAutomatenWare
    row% = .row
    col% = .col
    s$ = RTrim$(.TextMatrix(row%, 0))
    
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            

With frmEdit.lstEdit
    .Clear
    For i% = 1 To AnzLiefNamen%
        h$ = LiefNamen$(i% - 1)
        .AddItem h$
    Next i%
    
    .ListIndex = 0
    s$ = RTrim$(flxAutomatenWare.TextMatrix(row%, col%))
    If (s$ <> "") Then
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            If (s$ = .text) Then
                Exit For
            End If
        Next i%
    End If
    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxAutomatenWare.Left + flxAutomatenWare.ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxAutomatenWare.Top + flxAutomatenWare.RowPos(flxAutomatenWare.TopRow)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxAutomatenWare.ColWidth(col%)
        .Height = flxAutomatenWare.Height - flxAutomatenWare.RowPos(flxAutomatenWare.TopRow)
    End With
    With frmEdit.lstEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With
    
    frmEdit.Show 1
    
    With flxAutomatenWare
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then

        h$ = EditTxt$
        With flxAutomatenWare
            If (Left$(h$, 1) <> "-") Then
                .TextMatrix(row%, col%) = h$
                If (.col < .Cols - 2) Then .col = .col + 1
                If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
            End If
        End With
    End If

End With

Call DefErrPop
End Sub
                            

