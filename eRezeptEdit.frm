VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frm_eRezeptEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "Optionen"
   ClientHeight    =   11175
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   22095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   22095
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   8775
      Index           =   2
      Left            =   240
      ScaleHeight     =   8715
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   1800
      Width           =   9375
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Teilmengenabgabe (Spender-PZN)"
         Height          =   495
         Index           =   15
         Left            =   0
         TabIndex        =   207
         Top             =   7800
         Width           =   2175
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   15
         Left            =   2520
         TabIndex        =   206
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   7800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   15
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   205
         Text            =   "WWW9999"
         Top             =   7800
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Datum 
         Height          =   375
         Index           =   12
         Left            =   6720
         MaxLength       =   7
         TabIndex        =   87
         Text            =   "WWW9999"
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox txtZA_Uhrzeit 
         Height          =   375
         Index           =   10
         Left            =   7680
         MaxLength       =   7
         TabIndex        =   86
         Text            =   "WWW9999"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtZA_Datum 
         Height          =   375
         Index           =   10
         Left            =   6600
         MaxLength       =   7
         TabIndex        =   85
         Text            =   "WWW9999"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   14
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   84
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   7320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   13
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   83
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   12
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   82
         Text            =   "WWW9999"
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   11
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   81
         Text            =   "WWW9999"
         Top             =   5760
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   10
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   80
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   9
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   79
         Text            =   "WWW9999"
         Top             =   4800
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   8
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   78
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   7
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   77
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   6
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   76
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   5
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   75
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   6120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   4
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   74
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   3
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   73
         Text            =   "WWW9999"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   2
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   72
         Text            =   "WWW9999"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   1
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   71
         Text            =   "WWW9999"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   14
         Left            =   2520
         TabIndex        =   70
         Text            =   "Combo1"
         Top             =   7320
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   13
         Left            =   2640
         TabIndex        =   69
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   12
         Left            =   2640
         TabIndex        =   68
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   6360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   11
         Left            =   2760
         TabIndex        =   67
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   5880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   10
         Left            =   2400
         TabIndex        =   66
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   9
         Left            =   2400
         TabIndex        =   65
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   8
         Left            =   2400
         TabIndex        =   64
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   7
         Left            =   2400
         TabIndex        =   63
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   6
         Left            =   2400
         TabIndex        =   62
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   5
         Left            =   2400
         TabIndex        =   61
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   60
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   59
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   58
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   57
         Text            =   "Combo1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cboZA_Schluessel 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   56
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "ZuzahlungsStatus"
         Height          =   495
         Index           =   14
         Left            =   0
         TabIndex        =   55
         Top             =   7320
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "GruppeTarifkennzeichen"
         Height          =   495
         Index           =   13
         Left            =   120
         TabIndex        =   54
         Top             =   6840
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "GruppeFuerGenehmigung"
         Height          =   495
         Index           =   12
         Left            =   240
         TabIndex        =   53
         Top             =   6240
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "ZusaetzlicheAbgabeangaben"
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   52
         Top             =   5760
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "AbgabeNoctu"
         Height          =   495
         Index           =   10
         Left            =   0
         TabIndex        =   51
         Top             =   5160
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "EinzelimportierteFAM"
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   50
         Top             =   4680
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "KuenstlicheBefruchtung"
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   49
         Top             =   4200
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Ersatzverordnung"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   48
         Top             =   3720
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Wirkstoffverordnung"
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   47
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Wunscharzneimittel"
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   46
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Mehrkostenuebernahme"
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   45
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "ImportFAM"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   44
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "PreisguenstigesFAM"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Rabattvertragserfuellung"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   42
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox chkZA_Gruppe 
         Caption         =   "Markt"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Top             =   120
         Width           =   3615
      End
      Begin VB.TextBox txtZA_Freitext 
         Height          =   375
         Index           =   0
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblSpenderPzn 
         Height          =   375
         Left            =   4680
         TabIndex        =   208
         Top             =   8160
         Width           =   4575
      End
      Begin VB.Label lblchkZA_Gruppe 
         Caption         =   "AAAA"
         Height          =   375
         Index           =   0
         Left            =   7920
         TabIndex        =   5
         Top             =   4560
         Width           =   1095
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   5895
      Index           =   0
      Left            =   5280
      ScaleHeight     =   5835
      ScaleWidth      =   10995
      TabIndex        =   12
      Top             =   2160
      Width           =   11055
      Begin VB.TextBox txtPreise 
         Height          =   285
         Index           =   15
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "WWW9999"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   285
         Index           =   14
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   39
         Text            =   "WWW9999"
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   13
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   203
         Text            =   "WWW9999"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   11
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   30
         Text            =   "WWW9999"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   285
         Index           =   12
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   38
         Text            =   "WWW9999"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.ComboBox cboVertragsKz 
         Height          =   315
         Left            =   8760
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   8
         Left            =   9120
         MaxLength       =   30
         TabIndex        =   34
         Text            =   "WWW9999"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   7
         Left            =   8880
         MaxLength       =   30
         TabIndex        =   32
         Text            =   "WWW9999"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   6
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   26
         Text            =   "WWW9999"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   5
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   24
         Text            =   "WWW9999"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   4
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "WWW9999"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   3
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "WWW9999"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   2
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "WWW9999"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   1
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   16
         Text            =   "WWW9999"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   0
         Left            =   2760
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "WWW9999"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtPreise 
         Height          =   495
         Index           =   10
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   28
         Text            =   "WWW9999"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblPreise 
         Caption         =   "Nachtdienstgeb."
         Height          =   375
         Index           =   13
         Left            =   6480
         TabIndex        =   204
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label lblLieferengpass 
         AutoSize        =   -1  'True
         Caption         =   "( + 0.60 EUR Lieferengpasspauschale )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4320
         TabIndex        =   198
         Top             =   960
         Visible         =   0   'False
         Width           =   3360
      End
      Begin VB.Label lblPreise 
         Caption         =   "Beschaff.Kosten"
         Height          =   375
         Index           =   11
         Left            =   6600
         TabIndex        =   29
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "ChargenNr"
         Height          =   375
         Index           =   12
         Left            =   6600
         TabIndex        =   37
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Botendienst"
         Height          =   375
         Index           =   10
         Left            =   6600
         TabIndex        =   27
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "VertragsKz"
         Height          =   375
         Index           =   9
         Left            =   6720
         TabIndex        =   35
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Dosierung"
         Height          =   375
         Index           =   8
         Left            =   6720
         TabIndex        =   33
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Abgabedatum"
         Height          =   375
         Index           =   7
         Left            =   6480
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Eigenanteil"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   25
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Mehrkosten"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Zuzahlung"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Mwst"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Bruttopreis"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Gesamt-Brutto"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblPreise 
         Caption         =   "Gesamt-Zuz"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Index           =   4
      Left            =   840
      ScaleHeight     =   4155
      ScaleWidth      =   7515
      TabIndex        =   176
      Top             =   2640
      Width           =   7575
      Begin VB.CommandButton cmdDruckVerordnung 
         Caption         =   "Drucken (F6)"
         Height          =   450
         Left            =   1200
         TabIndex        =   201
         Top             =   360
         Width           =   2280
      End
      Begin SHDocVwCtl.WebBrowser wbVerordnung 
         Height          =   2535
         Left            =   1080
         TabIndex        =   177
         Top             =   960
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
      Begin nlCommandButton.nlCommand nlcmdDruckVerordnung 
         Height          =   375
         Left            =   3720
         TabIndex        =   202
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "nlCommand"
      End
      Begin VB.Label lblVerordnung 
         Caption         =   "Botendienst"
         Height          =   375
         Left            =   1080
         TabIndex        =   196
         Top             =   3600
         Width           =   4695
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   6015
      Index           =   6
      Left            =   2000
      ScaleHeight     =   5955
      ScaleWidth      =   10995
      TabIndex        =   180
      Top             =   0
      Width           =   11055
      Begin VB.CommandButton cmdQuittungErrneut 
         Caption         =   "Quittung erneut abrufen"
         Height          =   450
         Left            =   4560
         TabIndex        =   199
         Top             =   3720
         Width           =   2280
      End
      Begin SHDocVwCtl.WebBrowser wbXML 
         Height          =   2535
         Left            =   5640
         TabIndex        =   195
         Top             =   360
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
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   5
         Left            =   2640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   191
         Text            =   "eRezeptEdit.frx":0000
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   189
         Text            =   "eRezeptEdit.frx":0008
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   187
         Text            =   "eRezeptEdit.frx":0010
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   185
         Text            =   "WWW9999"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   1
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   183
         Text            =   "WWW9999"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtXML 
         Height          =   495
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   181
         Text            =   "WWW9999"
         Top             =   120
         Width           =   1575
      End
      Begin nlCommandButton.nlCommand nlcmdQuittungErrneut 
         Height          =   375
         Left            =   7680
         TabIndex        =   200
         Top             =   3720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "nlCommand"
      End
      Begin VB.Label lblXML 
         AutoSize        =   -1  'True
         Caption         =   "Dispense"
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   194
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label lblXML 
         AutoSize        =   -1  'True
         Caption         =   "Dispense"
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   193
         Top             =   3480
         Width           =   660
      End
      Begin VB.Label lblXML 
         Caption         =   "Abgabe"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   192
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label lblXML 
         Caption         =   "Dispense"
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   190
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblXML 
         Caption         =   "Verordnung"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   188
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblXML 
         Caption         =   "Secret"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   186
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblXML 
         Caption         =   "Access-Code"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   184
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblXML 
         Caption         =   "Task-Id"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   182
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Index           =   5
      Left            =   6960
      ScaleHeight     =   4155
      ScaleWidth      =   7515
      TabIndex        =   178
      Top             =   1440
      Width           =   7575
      Begin SHDocVwCtl.WebBrowser wbAbgabedaten 
         Height          =   2535
         Left            =   1080
         TabIndex        =   179
         Top             =   960
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
      Begin VB.Label lblAbgabedaten 
         Caption         =   "Botendienst"
         Height          =   375
         Left            =   1080
         TabIndex        =   197
         Top             =   3600
         Width           =   4575
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   5895
      Index           =   1
      Left            =   1440
      ScaleHeight     =   5835
      ScaleWidth      =   9915
      TabIndex        =   88
      Top             =   3120
      Width           =   9975
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   7
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   111
         Text            =   "WWW9999"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   7
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   112
         Text            =   "WWW9999"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.ComboBox cboVerordnungstyp 
         Height          =   315
         Left            =   6960
         TabIndex        =   175
         Text            =   "Combo1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   6
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   109
         Text            =   "WWW9999"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   6
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   108
         Text            =   "WWW9999"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   5
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   106
         Text            =   "WWW9999"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   5
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   105
         Text            =   "WWW9999"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   4
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   103
         Text            =   "WWW9999"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   4
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   102
         Text            =   "WWW9999"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   3
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   100
         Text            =   "WWW9999"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   3
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   99
         Text            =   "WWW9999"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   2
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   97
         Text            =   "WWW9999"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   96
         Text            =   "WWW9999"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   1
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   94
         Text            =   "WWW9999"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   93
         Text            =   "WWW9999"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtAbgabe 
         Height          =   495
         Index           =   0
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   91
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtVO 
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   90
         Text            =   "WWW9999"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblVO 
         Caption         =   "Anzahl Packungen"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   110
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Normgroesse"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   107
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Einheit"
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   104
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Packungsgroesse"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   101
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Darreichungsform"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   98
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Verordnungstext"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Pzn"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   92
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblVO 
         Caption         =   "Verordnungstyp"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   7935
      Index           =   3
      Left            =   12120
      ScaleHeight     =   7875
      ScaleWidth      =   9315
      TabIndex        =   113
      Top             =   480
      Width           =   9375
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   11
         Left            =   5040
         MaxLength       =   7
         TabIndex        =   171
         Text            =   "WWW9999"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   11
         Left            =   6120
         MaxLength       =   7
         TabIndex        =   172
         Text            =   "WWW9999"
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   10
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   166
         Text            =   "WWW9999"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   10
         Left            =   6000
         MaxLength       =   7
         TabIndex        =   167
         Text            =   "WWW9999"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   9
         Left            =   5040
         MaxLength       =   7
         TabIndex        =   161
         Text            =   "WWW9999"
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   9
         Left            =   6120
         MaxLength       =   7
         TabIndex        =   162
         Text            =   "WWW9999"
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   8
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   156
         Text            =   "WWW9999"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   8
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   157
         Text            =   "WWW9999"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   7
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   151
         Text            =   "WWW9999"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   7
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   152
         Text            =   "WWW9999"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   6
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   146
         Text            =   "WWW9999"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   6
         Left            =   6000
         MaxLength       =   7
         TabIndex        =   147
         Text            =   "WWW9999"
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   5
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   141
         Text            =   "WWW9999"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   5
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   142
         Text            =   "WWW9999"
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   4
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   136
         Text            =   "WWW9999"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   4
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   137
         Text            =   "WWW9999"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   3
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   131
         Text            =   "WWW9999"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   3
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   132
         Text            =   "WWW9999"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   2
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   126
         Text            =   "WWW9999"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   2
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   127
         Text            =   "WWW9999"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   1
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   121
         Text            =   "WWW9999"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   1
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   122
         Text            =   "WWW9999"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   0
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   118
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "AbweichungDAFO"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   114
         Top             =   120
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "ErgaenzungDAFORezeptur"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   119
         Top             =   600
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "Erg.GebrauchsanweisungRezeptur"
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   124
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "ErgaenzungDosieranweisung"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   129
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "ErgaenzungFehlenderHinweis"
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   134
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "AbweichungBezeichnungFAM "
         Height          =   495
         Index           =   5
         Left            =   0
         TabIndex        =   139
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "AbweichungBezeichnungWirkstoff"
         Height          =   495
         Index           =   6
         Left            =   0
         TabIndex        =   144
         Top             =   3240
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "AbweichungStaerke"
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   149
         Top             =   3720
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "Abw.Zus.setzungRezepturArtMenge"
         Height          =   495
         Index           =   8
         Left            =   0
         TabIndex        =   154
         Top             =   4200
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "AbweichungAbzugebendeMenge"
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   159
         Top             =   4680
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "Abw.RezepturmengeEntlassVO"
         Height          =   495
         Index           =   10
         Left            =   0
         TabIndex        =   164
         Top             =   5160
         Width           =   3615
      End
      Begin VB.CheckBox chkRAE_Art 
         Caption         =   "Sonstiges"
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   169
         Top             =   5760
         Width           =   3615
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   115
         Text            =   "Combo1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   120
         Text            =   "Combo1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   125
         Text            =   "Combo1"
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   130
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   135
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   5
         Left            =   2400
         TabIndex        =   140
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   6
         Left            =   2400
         TabIndex        =   145
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   7
         Left            =   2400
         TabIndex        =   150
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   8
         Left            =   2400
         TabIndex        =   155
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   9
         Left            =   2400
         TabIndex        =   160
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   4800
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   10
         Left            =   2400
         TabIndex        =   165
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   5280
         Width           =   1935
      End
      Begin VB.ComboBox cbokRAE_RueckspracheArzt 
         Height          =   315
         Index           =   11
         Left            =   2760
         TabIndex        =   170
         Tag             =   "0"
         Text            =   "Combo1"
         Top             =   5880
         Width           =   1935
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   1
         Left            =   7560
         MaxLength       =   7
         TabIndex        =   123
         Text            =   "WWW9999"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   2
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   128
         Text            =   "WWW9999"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   3
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   133
         Text            =   "WWW9999"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   4
         Left            =   7560
         MaxLength       =   30
         TabIndex        =   138
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   5
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   143
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   6
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   148
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   7
         Left            =   7440
         MaxLength       =   7
         TabIndex        =   153
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   8
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   158
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   9
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   163
         Text            =   "WWW9999"
         Top             =   4920
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   10
         Left            =   7440
         MaxLength       =   30
         TabIndex        =   168
         Tag             =   "0"
         Text            =   "WWW9999"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Dokumentation 
         Height          =   375
         Index           =   11
         Left            =   7680
         MaxLength       =   30
         TabIndex        =   173
         Text            =   "WWW9999"
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox txtRAE_Datum 
         Height          =   375
         Index           =   0
         Left            =   4800
         MaxLength       =   7
         TabIndex        =   116
         Text            =   "WWW9999"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtRAE_Uhrzeit 
         Height          =   375
         Index           =   0
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   117
         Text            =   "WWW9999"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblchkRAE_Art 
         Caption         =   "AAAA"
         Height          =   375
         Index           =   0
         Left            =   8040
         TabIndex        =   174
         Top             =   7080
         Width           =   1095
      End
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   10200
      Picture         =   "eRezeptEdit.frx":0018
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   10440
      Picture         =   "eRezeptEdit.frx":00C1
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   10680
      Picture         =   "eRezeptEdit.frx":0175
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   735
      Index           =   0
      Left            =   11160
      ScaleHeight     =   735
      ScaleWidth      =   2055
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   14280
      TabIndex        =   4
      Top             =   9360
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   12600
      TabIndex        =   3
      Top             =   9360
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   1365
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   2408
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
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
      TabCaption(0)   =   "&1 - Preise"
      TabPicture(0)   =   "eRezeptEdit.frx":022E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "&2 - Verordnung"
      TabPicture(1)   =   "eRezeptEdit.frx":024A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&3 - ZusatzAttribute"
      TabPicture(2)   =   "eRezeptEdit.frx":0266
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "&4 - RezeptÄnderungen"
      TabPicture(3)   =   "eRezeptEdit.frx":0282
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "&5 - eRezept"
      TabPicture(4)   =   "eRezeptEdit.frx":029E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "&6 - Abgabedaten "
      TabPicture(5)   =   "eRezeptEdit.frx":02BA
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "&7 - XML"
      TabPicture(6)   =   "eRezeptEdit.frx":02D6
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).ControlCount=   0
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   14640
      TabIndex        =   10
      Top             =   9960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   12840
      TabIndex        =   11
      Top             =   9960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frm_eRezeptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgRezApoDruckName$, OrgBtmRezDruckName$

Const PI = 3.14159265358979

Dim TabNamen$(8)
Dim TabEnabled%(8)
Dim AktTab%
Dim TabsPerRow%
Dim AnzTabs%

Dim iEditModus%

Dim ydiff%

Dim eRezept As TI_Back.eRezept_OP

Private Const DefErrModul = "EREZEPTEDIT.FRM"

Sub AbDatumEingeben()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbDatumEingeben")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ret%
Dim j%, l%, row%, ind%, aRow%
Dim s$, h$

'EditModus% = 4
'
'Load frmEdit
'
'With frmEdit
''    .Left = tabOptionen.Left + fmeOptionen(3).Left + flxOptionen1(2).Left + flxOptionen1(2).ColPos(1) + 45
'    .Left = picStammdatenBack(0).Left + flxOptionen1(2).Left + flxOptionen1(2).ColPos(1)
'    .Left = .Left + Me.Left + wpara.FrmBorderHeight
''    .Top = tabOptionen.Top + fmeOptionen(3).Top + flxOptionen1(2).Top + (flxOptionen1(2).row * flxOptionen1(2).RowHeight(0))
'    .Top = picStammdatenBack(0).Top + flxOptionen1(2).Top + (flxOptionen1(2).row - flxOptionen1(2).TopRow + 1) * flxOptionen1(2).RowHeight(0)
'    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
'    .Width = flxOptionen1(2).ColWidth(1)
'    .Height = flxOptionen1(2).RowHeight(0)
'End With
'With frmEdit.txtEdit
'    .Width = flxOptionen1(2).ColWidth(1)
'    .Left = 0
'    .Top = 0
'    h$ = flxOptionen1(2).TextMatrix(flxOptionen1(2).row, 1)
'    h$ = Left$(h$, 2) + Mid$(h$, 4, 2) + Mid$(h$, 7, 2)
'    .text = h$
'    .BackColor = vbWhite
'    .Visible = True
'End With
'
'frmEdit.Show 1
'
'If (EditErg%) Then
'    h$ = Trim(EditTxt$)
'    h$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "." + Mid$(h$, 5, 2)
'    If IsDate(h$) Then
'        flxOptionen1(2).TextMatrix(flxOptionen1(2).row, 1) = Format(CDate(h$), "dd.mm.yy")
'        With AbrechDatenRec
'            .MoveFirst
'            Do While Not .EOF
'                If AbrechDatenRec!Unique = flxOptionen1(2).row Then
'                    .Edit
'                    AbrechDatenRec!Datum = Trim(EditTxt$)
'                    .Update
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'        End With
'    End If
'End If

Call DefErrPop

End Sub

Private Sub cboZA_Schluessel_Click(Index As Integer)
    Call CheckLieferengpass
    Call CheckFreitext(Index)
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

If (wbXML.Visible) Then
    wbXML.Visible = False
Else
    Unload Me
End If

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
Dim i%, VERBAND%
Dim l&
Dim h$, h2$

If (ActiveControl.Name = cmdOk.Name) Or (ActiveControl.Name = nlcmdOk.Name) Then

    Dim iVal%, ind%
    Dim s As String
    
    eRezeptGespeichert = True
    
    With eRezept
        s = txtPreise(7).text
        .AbgabeDatum = CDate(Left(s, 2) + "." + Mid(s, 3, 2) + ".20" + Mid(s, 5, 2))
        .sGesamtBrutto = uFormat(xVal(txtPreise(1).text), "0.00")
        .sGesamtZuzahlung = uFormat(xVal(txtPreise(0).text), "0.00")
        .sBruttoPreis = uFormat(xVal(txtPreise(2).text), "0.00")
        .sMwst = uFormat(xVal(txtPreise(3).text), "0.00")
        .sZuzahlung = uFormat(xVal(txtPreise(4).text), "0.00")
        .sMehrkosten = uFormat(xVal(txtPreise(5).text), "0.00")
        .sEigenanteil = uFormat(xVal(txtPreise(6).text), "0.00")
        .Dosieranweisung = txtPreise(8).text
        .sBotendienst = uFormat(xVal(txtPreise(10).text), "0.00")
        .sBeschaffungsKosten = uFormat(xVal(txtPreise(11).text), "0.00")

'        .ChargenNr = txtPreise(12).text
        Dim sChargenNr As String
        sChargenNr = txtPreise(12).text
        For i = 14 To 15
            With txtPreise(i)
                If (.text <> "") Then
                    If (sChargenNr <> "") Then
                        sChargenNr = sChargenNr + "@"
                    End If
                    sChargenNr = sChargenNr + .text
                End If
            End With
        Next i
        .ChargenNr = sChargenNr
       
        iVal = 0
        With cboVertragsKz
            s = Trim(.text)
            If (s <> "") Then
                ind = InStr(s, "(")
                If (ind > 0) Then
                    s = Mid(s, ind + 1)
                    ind = InStr(s, ")")
                    If (ind > 0) Then
                        s = Strings.Left(s, ind - 1)
                        iVal = Val(s)
                    End If
                End If
            End If
        End With
        .Rechtsgrundlage = iVal
        
        .SpeicherAbgabedaten
        
        iVal = 0
        With cboVerordnungstyp
            s = Trim(.text)
            If (s <> "") Then
                ind = InStr(s, "(")
                If (ind > 0) Then
                    s = Mid(s, ind + 1)
                    ind = InStr(s, ")")
                    If (ind > 0) Then
                        s = Strings.Left(s, ind - 1)
                        iVal = Val(s)
                    End If
                End If
            End If
        End With
        .Abgabe.VerordnungsTyp = iVal
    
        .Abgabe.pzn = txtAbgabe(1).text
        .Abgabe.Verordnungstext = txtAbgabe(2).text
        .Abgabe.Darreichungsform = txtAbgabe(3).text
        .Abgabe.sPackungsgroesse = uFormat(xVal(txtAbgabe(4).text), "0.00")
        .Abgabe.einheit = txtAbgabe(5).text
        .Abgabe.Normgroesse = txtAbgabe(6).text
'        .Abgabe.Impfstoff = chkABG_Impfstoff.Checked

        .Anzahl = txtAbgabe(7).text
    
        .SpeicherAbgabeMedikation
    End With
   
    Dim TeilmengenAbgabeSpenderPzn As String
    i = 16
    If (eRezept.ZusatzAttribut(i).Freitext = "") And (chkZA_Gruppe(i - 1).Value) Then
        TeilmengenAbgabeSpenderPzn = txtZA_Freitext(i - 1).text
    End If
    
    For i = 1 To 16
        iVal = -1
        With chkZA_Gruppe(i - 1)
            If (.Value) Then
                If ((i >= 1) And (i <= 5)) Or (i = 15) Then
                    With cboZA_Schluessel(i - 1)
                        s = Trim(.text)
                        If (s <> "") Then
                            ind = InStr(s, "(")
                            If (ind > 0) Then
                                s = Mid(s, ind + 1)
                                ind = InStr(s, ")")
                                If (ind > 0) Then
                                    s = Strings.Left(s, ind - 1)
                                    iVal = Val(s)
                                End If
                            End If
                        End If
                    End With
                Else
                    iVal = 0    '1
                End If
            
                eRezept.ZusatzAttribut(i).Freitext = txtZA_Freitext(i - 1).text
                
                If (i = 11) Or (i = 13) Then
                    h = txtZA_Datum(i - 1).text
                    h2 = txtZA_Uhrzeit(i - 1).text
                    eRezept.ZusatzAttribut(i).Datum = CDate(Left(h, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5) + " " + Left(h2, 2) + ":" + Mid(h2, 3, 2))
                End If
            ElseIf (i = 16) Then
                eRezept.ZusatzAttribut(i).Freitext = ""
            End If
        End With
        eRezept.ZusatzAttribut(i).Schluessel = iVal
    Next
    eRezept.SpeicherZusatzAttribute


    For i = 1 To 12
        iVal = -1
        With chkRAE_Art(i - 1)
            If (.Value) Then
                With cbokRAE_RueckspracheArzt(i - 1)
                    s = Trim(.text)
                    If (s <> "") Then
                        ind = InStr(s, "(")
                        If (ind > 0) Then
                            s = Mid(s, ind + 1)
                            ind = InStr(s, ")")
                            If (ind > 0) Then
                                s = Strings.Left(s, ind - 1)
                                iVal = Val(s)
                                
                                If (txtRAE_Datum(i - 1).Visible) Then
                                    h = txtRAE_Datum(i - 1).text
                                    h2 = txtRAE_Uhrzeit(i - 1).text
                                    eRezept.RezeptAenderung(i).DatumAenderung = CDate(Left(h, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5) + " " + Left(h2, 2) + ":" + Mid(h2, 3, 2))
                                End If
                                If (txtRAE_Dokumentation(i - 1).Visible) Then
                                    h = Trim(txtRAE_Dokumentation(i - 1).text)
                                    eRezept.RezeptAenderung(i).Dokumentation = h
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        End With
        eRezept.RezeptAenderung(i).RueckspracheArzt = iVal
    Next
    eRezept.SpeicherRezeptAenderungen

    If (TeilmengenAbgabeSpenderPzn <> "") Then
        With eRezept.Abgabe
            If (MessageBox("Soll die gespeicherte Abgabe-PZN" + vbCrLf + vbCrLf + .pzn + "  " + .Verordnungstext + " " + .sPackungsgroesse + " " + .Darreichungsform + vbCrLf + vbCrLf + "rückgebucht werden?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Teilmengenabgabe") = vbYes) Then
                Call WuBuchen(.pzn, eRezept.Anzahl)
            End If
        End With
        If (MessageBox("Soll die neu eingetragene Spender-PZN" + vbCrLf + vbCrLf + TeilmengenAbgabeSpenderPzn + "  " + lblSpenderPzn.Caption + vbCrLf + vbCrLf + "abgebucht werden?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Teilmengenabgabe") = vbYes) Then
            Call WuBuchen(TeilmengenAbgabeSpenderPzn, -eRezept.Anzahl)
        End If
    End If

'    Call AuslesenFlexTaetigkeiten
'
'    h$ = Right$(Space$(7) + Trim(txtOptionen0(0).text), 7)
'    If (h$ <> OrgRezApoNr$) Then
'        OrgRezApoNr$ = h$
'        RezApoNr$ = h$
'        l& = WritePrivateProfileString("Rezeptkontrolle", "InstitutsKz", h$, INI_DATEI)
'    End If
'
'    h$ = Trim(txtOptionen0(1).text)
'    If (h$ <> OrgRezApoDruckName$) Then
'        RezApoDruckName$ = h$
'        l& = WritePrivateProfileString("Rezeptkontrolle", "RezeptText", h$, INI_DATEI)
'    End If
'
'    h$ = Trim(txtOptionen0(2).text)
'    If (h$ <> OrgBtmRezDruckName$) Then
'        BtmRezDruckName$ = h$
'        l& = WritePrivateProfileString("Rezeptkontrolle", "BtmRezeptText", h$, INI_DATEI)
'    End If
'
''    VmRabattFaktor# = Val(txtOptionen0(1).text)
'    VmRabattFaktor# = 100# / (100# - Val(txtOptionen0(3).text))
'
'
'    OrgBundesland% = Val(flxOptionen1(0).TextMatrix(flxOptionen1(0).row, 1))
'
'    VERBAND% = FileOpen("verbandm.dat", "RW", "B")
'
'    h$ = MKI(OrgBundesland%)
'    Seek #VERBAND%, 7
'    Put #VERBAND%, , h$
'
'    h$ = Format(VmRabattFaktor#, "0.0000")
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    Seek #VERBAND%, 9
'    Put #VERBAND%, , h$
'    Close #VERBAND%
'
'    h$ = "N"
'    RezepturMitFaktor% = False
'    If (chkOptionen0(0).Value) Then
'        RezepturMitFaktor% = True
'        h$ = "J"
'    End If
'    l& = WritePrivateProfileString("Rezeptkontrolle", "RezepturMitFaktor", h$, INI_DATEI)
'
'    h$ = "N"
'    BtmAlsZeile% = False
'    If (chkOptionen0(1).Value) Then
'        BtmAlsZeile% = True
'        h$ = "J"
'    End If
'    l& = WritePrivateProfileString("Rezeptkontrolle", "BtmAlsZeile", h$, INI_DATEI)
'
'    h$ = "N"
'    RezepturDruck% = False
'    If (chkOptionen0(2).Value) Then
'        RezepturDruck% = True
'        h$ = "J"
'    End If
'    l& = WritePrivateProfileString("Rezeptkontrolle", "RezepturDruck", h$, INI_DATEI)
'
'    h$ = "N"
'    AvpTeilnahme% = False
'    If (chkOptionen0(3).Value) Then
'        AvpTeilnahme% = True
'        h$ = "J"
'    End If
'    l& = WritePrivateProfileString("Rezeptkontrolle", "AvpTeilnahme", h$, INI_DATEI)
'
'    For i% = 0 To UBound(ParenteralPzn)
'        h$ = ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00")
'        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i), h$, INI_DATEI)
'    Next i%
'
'    For i% = 0 To UBound(ParEnteralAufschlag)
'        h$ = Format(ParEnteralAufschlag(i), "0.00")
'        l& = WritePrivateProfileString("Parenteral", "Aufschlag" + CStr(i), h$, INI_DATEI)
'    Next i%
'
'    frmAction.mnuDateiInd(6).Enabled = AvpTeilnahme%
'    frmAction.cmdDatei(6).Enabled = AvpTeilnahme%
'
'    OptionenNeu% = True
'    Call AbrechMonatErmitteln
    Unload Me
'ElseIf (ActiveControl.Name = flxOptionen1(0).Name) Then
'    If (ActiveControl.index = 1) Then
'        Call EditOptionenLstMulti
'    ElseIf (ActiveControl.index = 2) Then
'        Call AbDatumEingeben
'    ElseIf (ActiveControl.index = 3) Then
'        Call EditOptionenTxt(ActiveControl.index)
'    ElseIf (ActiveControl.index = 4) Then
'        Call EditOptionenTxt(ActiveControl.index)
'    ElseIf (ActiveControl.index = 5) Then
'        Call EditOptionenTxt(ActiveControl.index)
'    End If
End If

Call DefErrPop
End Sub

Private Sub flxOptionen1_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'With flxOptionen1(index)
'    If (index = 0) Then
'        .col = 0
'        .ColSel = .Cols - 1
'    End If
'    .HighLight = flexHighlightAlways
'End With

Call DefErrPop
End Sub

Private Sub cmdQuittungErrneut_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdQuittungErrneut_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (FD_OP.TI_QuittungErneut(eRezept.TaskId, eRezept.Secret)) Then
    Call MessageBox("'Quittung Erneut abrufen' ERFOLGREICH !", vbInformation)
Else
    Call MessageBox("Probleme bei 'Quittung Erneut abrufen' !", vbCritical)
End If
    
Call DefErrPop
End Sub

Private Sub flxOptionen1_LostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'With flxOptionen1(index)
'    .HighLight = flexHighlightNever
'End With

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
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim iAdd%, iAdd2%
Dim h$, h2$, h3$, FormStr$
Dim c As Control


eRezeptGespeichert = 0

iEditModus = 1

Call wpara.InitFont(Me)

Dim sTaskId$
If (eRezeptTaskId <> "") Then
    sTaskId = eRezeptTaskId
Else
    With frm_eRezepte.flxarbeit(0)
        sTaskId = .TextMatrix(.row, 2)
    End With
End If
Set eRezept = New TI_Back.eRezept_OP
Call eRezept.New2(sTaskId)

Call RefreshTabControls

With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
'    .Top = tabOptionen.Top + tabOptionen.Height + 300
    .Top = picStammdatenBack(0).Top + picStammdatenBack(0).Height + 300
End With
With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Top = cmdOk.Top
End With

With cmdQuittungErrneut
    .Width = TextWidth(.Caption) + 1200
    .Height = wpara.ButtonY
    .Top = lblXML(6).Top + lblXML(6).Height + 300
    .Left = lblXML(6).Left
End With
With cmdDruckVerordnung
    .Width = TextWidth(.Caption) + 1200
    .Height = wpara.ButtonY
    .Top = wbVerordnung.Top + wbVerordnung.Height + 90
    .Left = wbVerordnung.Left + wbVerordnung.Width - .Width
End With

'Me.Width = tabOptionen.Left + tabOptionen.Width + 2 * wpara.LinksX
Me.Width = picStammdatenBack(0).Left + picStammdatenBack(0).Width + 2 * wpara.LinksX
Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    Height = Height + iAdd2
'    Width = Width + iAdd2 + 600
    
    With nlcmdOk
        .Init
'        .Left = (Me.ScaleWidth - 2 * .Width - 300)
'        .Top = tabProfil.Top + tabProfil.Height + iAdd + 600
        .Top = picStammdatenBack(0).Top + picStammdatenBack(0).Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False
    
    With nlcmdQuittungErrneut
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = cmdQuittungErrneut.Top
        .Left = cmdQuittungErrneut.Left
        .Caption = cmdQuittungErrneut.Caption
        .TabIndex = cmdQuittungErrneut.TabIndex
        .Enabled = cmdQuittungErrneut.Enabled
        .Default = cmdQuittungErrneut.Default
        .Cancel = cmdQuittungErrneut.Cancel
        
        .Width = cmdQuittungErrneut.Width
        .Visible = True
    End With
    cmdQuittungErrneut.Visible = False
    
    With nlcmdDruckVerordnung
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = cmdDruckVerordnung.Top
        .Left = cmdDruckVerordnung.Left
        .Caption = cmdDruckVerordnung.Caption
        .TabIndex = cmdDruckVerordnung.TabIndex
        .Enabled = cmdDruckVerordnung.Enabled
        .Default = cmdDruckVerordnung.Default
        .Cancel = cmdDruckVerordnung.Cancel
        
        .Width = cmdDruckVerordnung.Width
        .Visible = True
    End With
    cmdDruckVerordnung.Visible = False
    
'    Me.Width = nlcmdImport(0).Left + nlcmdImport(0).Width + 600

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
'    With flxAbglPartner
'        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'    End With

    On Error Resume Next
    For Each c In Controls
'        If (c.Container Is Me) Then
            If (c.tag <> "0") Then
                If (TypeOf c Is Label) Then
                    c.BackStyle = 0 'duchsichtig
                ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Or (TypeOf c Is ListBox) Then
                    If (TypeOf c Is ComboBox) Then
                        Call wpara.ControlBorderless(c)
                    ElseIf (c.Appearance = 1) Then
                        Call wpara.ControlBorderless(c, 2, 2)
                    Else
                        Call wpara.ControlBorderless(c, 1, 1)
                    End If
    
                    If (c.Enabled) Then
                        c.BackColor = vbWhite
                    Else
                        c.BackColor = Me.BackColor
                    End If
    
    '                If (c.Visible) Then
                        With c.Container
                            .ForeColor = RGB(180, 180, 180) ' vbWhite
                            .FillStyle = vbSolid
                            .FillColor = c.BackColor
    
                            RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                        End With
    '                End If
                ElseIf (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
                    With c
'                        .BackColor = GetPixel(.Container.hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                        .BackColor = GetPixel(picStammdatenBack(0).hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                        .Height = 0
                        .Width = .Height * 3 / 4
                    End With
                    If (c.Name = "chkZA_Gruppe") Then
                        If (c.Index > 0) Then
                            Load lblchkZA_Gruppe(c.Index)
                        End If
                        With lblchkZA_Gruppe(c.Index)
                            .BackStyle = 0 'duchsichtig
                            .Caption = c.Caption
                            .Left = c.Left + c.Width + 60
                            .Top = c.Top
                            .Width = TextWidth(.Caption) + 90
                            .TabIndex = c.TabIndex
                            .Visible = True
                        End With
                    ElseIf (c.Name = "chkRAE_Art") Then
                        If (c.Index > 0) Then
                            Load lblchkRAE_Art(c.Index)
                        End If
                        With lblchkRAE_Art(c.Index)
                            .BackStyle = 0 'duchsichtig
                            .Caption = c.Caption
                            .Left = c.Left + c.Width + 60
                            .Top = c.Top
                            .Width = TextWidth(.Caption) + 90
                            .TabIndex = c.TabIndex
                            .Visible = True
                        End With
                    End If
                ElseIf (TypeOf c Is MSFlexGrid) Then
                    With c
'                        .Left = (c.Container.Width - .Width) / 2
                        .Left = (picStammdatenBack(0).Width - .Width) / 2
                        With c.Container
                            .ForeColor = RGB(180, 180, 180) ' vbWhite
                            .FillStyle = vbSolid
                            .FillColor = c.BackColor
                        End With
                        RoundRect c.Container.hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
                    End With
                End If
'            End If
        End If
    Next
    On Error GoTo DefErr
    
    
    With Me
        ind = 0
        If (TabsPerRow < AnzTabs) Then
            ind = TabsPerRow
        End If
        
        .ForeColor = RGB(180, 180, 180) ' vbWhite
    
        .FillStyle = vbSolid
        .FillColor = RGB(232, 217, 172)
        .FillColor = RGB(200, 200, 200)
        RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX - 1, picTab(ind).Top / Screen.TwipsPerPixelY - 1, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX + 1, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 + 10, 20, 20
    
        .FillColor = RGB(200, 200, 200)
    '    RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX - 1, picTab(0).Top / Screen.TwipsPerPixelY - 1, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX + 1, picTab(0).Top / Screen.TwipsPerPixelY + 20, 10, 10
    '    Me.Line (picTab(0).Left + 15, picTab(0).Top + 150)-(picStammdatenBack(0).Left + picStammdatenBack(0).Width - 30, picTab(0).Top + 600), .FillColor, BF
    
        .FillColor = RGB(232, 217, 172)
        .ForeColor = .FillColor
        RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 - 11, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 + 9, 20, 20
    End With
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

'tabOptionen.Tab = 0
TabEnabled(0) = True
Call picTab_Click(2)

'Call InitAnimation

Call DefErrPop
End Sub

'txtOptionen0(1).text = String(38, "A")
'txtOptionen0(2).text = String(38, "A")
'
'On Error Resume Next
'For Each c In Controls
'    If (TypeOf c Is TextBox) Then
'        c.Width = TextWidth(c.text) + 90
'        c.text = ""
'    End If
'Next
'On Error GoTo DefErr
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 0
'
'lblOptionen0(0).Left = wpara.LinksX
'lblOptionen0(0).Top = 2 * wpara.TitelY
'txtOptionen0(0).Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 300
'txtOptionen0(0).Top = lblOptionen0(0).Top + (lblOptionen0(0).Height - txtOptionen0(0).Height) / 2
'
'For i% = 1 To 3
'    lblOptionen0(i%).Left = lblOptionen0(0).Left
'    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
'    txtOptionen0(i%).Left = txtOptionen0(0).Left
'    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
'Next i%
'
'For i% = 0 To 3
'    With chkOptionen0(i%)
'        .Left = lblOptionen0(0).Left
'        If (i% = 0) Then
'            .Top = lblOptionen0(3).Top + lblOptionen0(3).Height + 600
'        Else
'            .Top = chkOptionen0(i% - 1).Top + chkOptionen0(i% - 1).Height + 150
'        End If
'    End With
'Next i%
'
'txtOptionen0(0).text = OrgRezApoNr$
'OrgRezApoDruckName$ = RezApoDruckName$
'txtOptionen0(1).text = Trim(Left$(RezApoDruckName$ + Space$(38), 38))
'OrgBtmRezDruckName$ = BtmRezDruckName$
'txtOptionen0(2).text = Trim(Left$(BtmRezDruckName$ + Space$(50), 50))
'
'txtOptionen0(3).text = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")
'
'chkOptionen0(0).Value = Abs(RezepturMitFaktor%)
'chkOptionen0(1).Value = Abs(BtmAlsZeile%)
'chkOptionen0(2).Value = Abs(RezepturDruck%)
'chkOptionen0(3).Value = Abs(AvpTeilnahme%)
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 1
'Call ActProgram.FlxOptionenBefuellen
'With flxOptionen1(0)
'    Breite1% = 0
'    For i% = 0 To (.Rows - 1)
'        Breite2% = TextWidth(.TextMatrix(i%, 0))
'        If (Breite2% > Breite1%) Then Breite1% = Breite2%
'    Next i%
'    .ColWidth(0) = Breite1% + 150
'    .ColWidth(1) = TextWidth("00000")
'    .ColWidth(2) = wpara.FrmScrollHeight
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
'    .Height = .RowHeight(0) * 11 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 2
'With flxOptionen1(1)
'    .Cols = 3
'    .Rows = 2
'    .FixedRows = 1
'    .Rows = 1
'
'    .FormatString = "^Tätigkeit|^Personal|^ "
'
'    .ColWidth(0) = TextWidth(String(20, "A")) + 150
'    .ColWidth(1) = TextWidth(String(30, "A"))
'    .ColWidth(2) = 0
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
'    .Height = .RowHeight(0) * 11 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'
'    h$ = "Rezeptspeicher" + vbTab
'
'    For i% = 0 To (AnzTaetigkeiten% - 1)
''        h$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
''        h$ = h$ + vbTab
'        h2$ = ""
'        For k% = 0 To 79
'            If (Taetigkeiten(i%).pers(k%) > 0) Then
'                h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
'            Else
'                Exit For
'            End If
'        Next k%
'        If (k% = 1) Then
'            ind% = Taetigkeiten(i%).pers(0)
'            h3$ = RTrim$(para.Personal(ind%))
'        ElseIf (k% > 1) Then
'            h3$ = "mehrere (" + Mid$(Str$(k%), 2) + ")"
'        End If
'        h$ = h$ + h3$ + vbTab + h2$
'    Next i%
'
'    .AddItem h$
'    .row = 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 3
'With flxOptionen1(2)
'    .Rows = 13
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 1
'    .TextMatrix(0, 0) = "Monat"
'    .TextMatrix(0, 1) = "Datum"
'    .ColWidth(0) = TextWidth(String$(10, "A"))
'    .ColWidth(1) = TextWidth(String$(10, "0"))
'    AbrechDatenRec.MoveFirst
'    For i% = 1 To 12
'        .TextMatrix(i%, 0) = Format(CDate("01." + CStr(i%) + ".2002"), "MMMM")
'        .TextMatrix(i%, 1) = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
'        AbrechDatenRec.MoveNext
'    Next i%
'
'    .Width = .ColWidth(0) + .ColWidth(1) + 90
'    .Height = .RowHeight(0) * 13 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 4
'With flxOptionen1(3)
'    .Rows = 11
'    .Cols = 5
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<SonderPzn|<KK-Bez|<KassenNr|<Status|<GültigBis"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(13, "X"))
'    .ColWidth(2) = TextWidth(String$(10, "9"))
'    .ColWidth(3) = TextWidth(String$(10, "9"))
'    .ColWidth(4) = TextWidth(String$(10, "9"))
''    .ColWidth(5) = wpara.FrmScrollHeight
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    For i% = 1 To AnzSonderBelege%
'        .TextMatrix(i%, 0) = SonderBelege(i% - 1).pzn
'        .TextMatrix(i%, 1) = SonderBelege(i% - 1).KkBez
'        .TextMatrix(i%, 2) = SonderBelege(i% - 1).KassenId
'        .TextMatrix(i%, 3) = SonderBelege(i% - 1).Status
'        .TextMatrix(i%, 4) = SonderBelege(i% - 1).GültigBis
'    Next i%
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 5
'With flxOptionen1(4)
'    .Rows = 9
'    .Cols = 3
'    .FixedRows = 1
'    .FixedCols = 2
'
'    .FormatString = "<SonderPzn|<Bezeichnung|>ArbeitsPreis (EUR)"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(42, "X"))
'    .ColWidth(2) = TextWidth(String$(17, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    .Rows = .FixedRows
'    For i% = 0 To UBound(ParenteralPzn)
'        h = ParenteralPzn(i) + vbTab + ParenteralTxt(i) + vbTab + Format(ParenteralPreis(i), "0.00")
'        .AddItem h$
'    Next i%
'
'    .row = .FixedRows
'    .col = .Cols - 1
'    .ColSel = .col
'End With
'
'With lblOptionen5(0)
'    .Left = wpara.LinksX
'    .Top = flxOptionen1(4).Top + flxOptionen1(4).Height + 450
'    txtOptionen5(0).Left = .Left + .Width + 150
'    txtOptionen5(0).Top = .Top + (.Height - txtOptionen0(0).Height) / 2
'End With
'For i% = 1 To 1
'    lblOptionen5(i%).Left = lblOptionen5(0).Left
'    lblOptionen5(i%).Top = lblOptionen5(i% - 1).Top + lblOptionen0(i% - 1).Height + 150
'    txtOptionen5(i%).Left = txtOptionen5(0).Left
'    txtOptionen5(i%).Top = lblOptionen5(i%).Top + (lblOptionen5(i%).Height - txtOptionen5(i%).Height) / 2
'Next i%
'For i = 0 To 1
'    txtOptionen5(i).text = Format(ParEnteralAufschlag(i), "0.00")
'Next i
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 6
'With flxOptionen1(5)
'    .Rows = 11
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<AbgabeArt|>Preis"
'
'    .ColWidth(0) = TextWidth(String$(30, "X"))
'    .ColWidth(1) = TextWidth(String$(10, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    Call ActProgram.LadeOptionenAbgabeKosten
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
''''''''''''''''''''''''''''''''''''
'
'
'Font.Name = wpara.FontName(1)
'Font.Size = wpara.FontSize(1)
'
'With fmeOptionen(0)
'    .Left = wpara.LinksX
'    .Top = 3 * wpara.TitelY
''    .Width = txtOptionen0(1).Left + txtOptionen0(1).Width + 900
''    .Width = flxOptionen1(1).Left + flxOptionen1(1).Width + 1800
'    .Width = flxOptionen1(4).Left + flxOptionen1(4).Width + 900
'
'    Hoehe1% = chkOptionen0(3).Top + chkOptionen0(3).Height
'    Hoehe2% = flxOptionen1(2).Top + flxOptionen1(2).Height
'    If (Hoehe2% > Hoehe1%) Then
'        Hoehe1% = Hoehe2%
'    End If
'    .Height = Hoehe1% + 300
'End With
'For i% = 1 To 6
'    With fmeOptionen(i%)
'        .Left = fmeOptionen(0).Left
'        .Top = fmeOptionen(0).Top
'        .Width = fmeOptionen(0).Width
'        .Height = fmeOptionen(0).Height
'    End With
'Next i%
'
'With tabOptionen
'    .Left = wpara.LinksX
'    .Top = wpara.TitelY
'    .Width = fmeOptionen(0).Left + fmeOptionen(0).Width + wpara.LinksX
'    .Height = fmeOptionen(0).Top + fmeOptionen(0).Height + wpara.TitelY
'End With
'
'
'
'cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
'cmdEsc.Top = cmdOk.Top
'
'Me.Width = tabOptionen.Width + 2 * wpara.LinksX
'
'cmdOk.Width = wpara.ButtonX
'cmdOk.Height = wpara.ButtonY
'cmdEsc.Width = wpara.ButtonX
'cmdEsc.Height = wpara.ButtonY
'cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
'cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
'
'Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight
'
'Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
'If (Breite1% < 0) Then Breite1% = 0
'Me.Left = Breite1%
'Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
'If (Hoehe1% < 0) Then Hoehe1% = 0
'Me.Top = Hoehe1%
'
'tabOptionen.Tab = 0
'Call TabDisable
'Call TabEnable(tabOptionen.Tab)
'
'Call DefErrPop
'End Sub

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

If (Shift And vbAltMask) And (KeyCode >= 49) And (KeyCode <= 57) Then
    Call picTab_Click(KeyCode - 49)
End If

If (para.Newline) Then
    If (KeyCode = vbKeyF6) Then
        nlcmdDruckVerordnung.Value = True
    End If
Else
    If (KeyCode = vbKeyF6) Then
        cmdDruckVerordnung.Value = True
    End If
End If


'If (KeyCode = vbKeyF2) Then
'    cmdF2.Value = True
'End If

Call DefErrPop
End Sub

'Private Sub tabOptionen_Click(PreviousTab As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("tabOptionen_Click")
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
'If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub
'
'Call TabDisable
'Call TabEnable(tabOptionen.Tab)
'
'Call DefErrPop
'End Sub

Private Sub flxOptionen1_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_DblClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

'Sub TabDisable()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TabDisable")
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
'Dim i%
'
'For i% = 0 To 6
'    fmeOptionen(i%).Visible = False
'Next i%
'
'Call DefErrPop
'End Sub
'
'Sub TabEnable(hTab%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TabEnable")
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
'Dim i%
'
'fmeOptionen(hTab%).Visible = True
'
'If (hTab% = 0) Then
'    If (txtOptionen0(0).Visible) Then txtOptionen0(0).SetFocus
'ElseIf (hTab% = 1) Then
'    flxOptionen1(0).SetFocus
'ElseIf (hTab% = 2) Then
'    flxOptionen1(1).col = 1
'    flxOptionen1(1).SetFocus
'ElseIf (hTab% = 3) Then
'    flxOptionen1(2).col = 1
'    flxOptionen1(2).SetFocus
'ElseIf (hTab% = 4) Then
'    flxOptionen1(3).col = 0
'    flxOptionen1(3).SetFocus
'ElseIf (hTab% = 5) Then
'    flxOptionen1(4).col = 2
'    flxOptionen1(4).SetFocus
'ElseIf (hTab% = 6) Then
'    flxOptionen1(5).col = 0
'    flxOptionen1(5).SetFocus
'End If
'
'Call DefErrPop
'End Sub

Private Sub txtOptionen0_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

'With txtOptionen0(index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
'    .SelStart = 0
'    .SelLength = Len(.text)
'End With

Call DefErrPop
End Sub

Private Sub txtOptionen0_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Index <> 1) And (Index <> 2) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((Index <> 2) Or (Chr$(KeyAscii) <> ".")) Then
        Beep
        KeyAscii = 0
    End If
End If

Call DefErrPop
End Sub

Private Sub txtOptionen5_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen5_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

'With txtOptionen5(index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
'    .SelStart = 0
'    .SelLength = Len(.text)
'End With

Call DefErrPop
End Sub

Private Sub txtOptionen5_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen5_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
    Beep
    KeyAscii = 0
End If

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

'hTab% = tabOptionen.Tab
''row% = 1
'col% = 1
'
'With flxOptionen1(1)
'    row = .row
'    aRow% = .row
'    .row = 0
'    .CellFontBold = True
'    .row = aRow%
'End With
'
'With frmEdit.lstMultiEdit
'    .Clear
'    .AddItem "(keiner)"
'    For i% = 1 To 80
'        h$ = para.Personal(i%)
'        .AddItem h$
'    Next i%
'
'    For i% = 0 To (.ListCount - 1)
'        .Selected(i%) = False
'    Next i%
'
'
'    Load frmEdit
'
'     .ListIndex = 0
'
'     BetrLief$ = LTrim$(RTrim$(flxOptionen1(1).TextMatrix(row%, 2)))
'
'     For i% = 0 To 19
'         If (BetrLief$ = "") Then Exit For
'
'         ind% = InStr(BetrLief$, ",")
'         If (ind% > 0) Then
'             Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
'             BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
'         Else
'             Lief2$ = BetrLief$
'             BetrLief$ = ""
'         End If
'
'         If (Lief2$ <> "") Then
'            ind% = Val(Lief2$)
'            .Selected(ind%) = True
'         End If
'     Next i%
'
'    With frmEdit
''        .Left = tabOptionen.Left + fmeOptionen(2).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col%) + 45
'        .Left = picStammdatenBack(0).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col)
'        .Left = .Left + Me.Left + wpara.FrmBorderHeight
''        .Top = tabOptionen.Top + fmeOptionen(2).Top + flxOptionen1(1).Top + flxOptionen1(1).RowHeight(0)
'        .Top = picStammdatenBack(0).Top + flxOptionen1(1).Top + (row% - flxOptionen1(1).TopRow + 1) * flxOptionen1(1).RowHeight(0)
'        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
'        .Width = flxOptionen1(1).ColWidth(col%)
'        .Height = flxOptionen1(1).Height - flxOptionen1(1).RowHeight(0)
'    End With
'    With frmEdit.lstMultiEdit
'        .Height = frmEdit.ScaleHeight
'        frmEdit.Height = .Height
'        .Width = frmEdit.ScaleWidth
'        .Left = 0
'        .Top = 0
'
'        .Visible = True
'    End With
'
'
'    frmEdit.Show 1
'
'    With flxOptionen1(1)
'        aRow% = .row
'        .row = 0
'        .CellFontBold = False
'        .row = aRow%
'    End With
'
'
'    If (EditErg%) Then
'
'        flxOptionen1(1).TextMatrix(row%, col% + 1) = EditTxt$
'
'        h$ = ""
'        If (EditAnzGefunden% = 0) Then
'            h$ = ""
'        ElseIf (EditAnzGefunden% = 1) Then
'            ind% = EditGef%(0)
'            h$ = RTrim$(para.Personal(ind%))
'        Else
'            h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
'        End If
'
'        With flxOptionen1(1)
'            .TextMatrix(row%, col%) = h$
'            If (.col < .Cols - 2) Then .col = .col + 1
'            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
'        End With
'
'    End If
'
'End With

Call DefErrPop
End Sub
                           
Private Sub picTab_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picTab_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (tabStammdaten.Visible = False) Then Call DefErrPop: Exit Sub

If (TabEnabled(0)) And (TabEnabled(Index + 1)) Then
    Call TabDisable
    Call TabEnable(Index)
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

AktTab = -1
For i% = 0 To 6
    Call PaintTab(i)
Next i
'cmdImport(0).Visible = False
'cmdImport(1).Visible = False
'nlcmdImport(0).Visible = False
'nlcmdImport(1).Visible = False

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

AktTab = hTab
Call PaintTab(hTab)

If (Me.Visible) Then
    If (hTab% = 0) Then
        If (chkZA_Gruppe(0).Visible) Then chkZA_Gruppe(0).SetFocus
'    ElseIf (hTab% = 1) Then
'        flxOptionen1(0).SetFocus
'    ElseIf (hTab% = 2) Then
'        flxOptionen1(1).col = 1
'        flxOptionen1(1).SetFocus
'    ElseIf (hTab% = 3) Then
'        flxOptionen1(2).col = 1
'        flxOptionen1(2).SetFocus
'    ElseIf (hTab% = 4) Then
'        flxOptionen1(3).col = 0
'        flxOptionen1(3).SetFocus
'    ElseIf (hTab% = 5) Then
'        flxOptionen1(4).col = 2
'        flxOptionen1(4).SetFocus
'    ElseIf (hTab% = 6) Then
'        flxOptionen1(5).col = 0
'        flxOptionen1(5).SetFocus
    End If
End If

Call DefErrPop
End Sub

'Private Sub lblchkOptionen0_Click(index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("lblchkOptionen0_Click")
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
'
'With chkOptionen0(index)
'    If (.Enabled) Then
'        If (.Value) Then
'            .Value = 0
'        Else
'            .Value = 1
'        End If
'        .SetFocus
'    End If
'End With
'
'Call DefErrPop
'End Sub
'
'Private Sub chkOptionen0_GotFocus(index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("chkOptionen0_GotFocus")
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
'
'Call nlCheckBox(chkOptionen0(index).Name, index)
'
'Call DefErrPop
'End Sub
'
'Private Sub chkOptionen0_LostFocus(index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("chkOptionen0_LostFocus")
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
'
'Call nlCheckBox(chkOptionen0(index).Name, index, 0)
'
'Call DefErrPop
'End Sub
'
'Sub nlCheckBox(sCheckBox$, index As Integer, Optional GotFocus% = True)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("nlCheckBox")
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
'Dim ok%
'Dim Such$
'Dim c As Object
'
'Such = "lbl" + sCheckBox
'
'On Error Resume Next
'For Each c In Controls
'    If (c.Name = Such) Then
'        ok = True
'        If (index >= 0) Then
'            ok = (c.index = index)
'        End If
'        If (ok) Then
'            If (GotFocus) Then
''                c.Font.underline = True
''                c.ForeColor = vbHighlight
'                c.BackStyle = 1
'                c.BackColor = vbHighlight
'                c.ForeColor = vbWhite
'            Else
''                c.Font.underline = 0
''                c.ForeColor = vbBlack
'                c.BackStyle = 0
'                c.BackColor = vbHighlight
'                c.ForeColor = vbBlack
'            End If
'        End If
'    End If
'Next
'On Error GoTo DefErr
'
'Call DefErrPop
'End Sub


Sub RefreshTabControls()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshTabControls")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, TxtHe%, LblHe%, wi%, MaxWi%, iAdd%, x%, y%, spBreite%, iRows%, RowsNeeded%, Breite1%, Breite2%, Hoehe1%, Hoehe2%, xpos%, iVal%
Dim OrgTab%, RowHe%, AnzZe%, ind%, ind2%, x2%
Dim von&
Dim h$, h2$, h3$, FormStr$, s$
Dim c As Control

Font.Name = wpara.FontName(0)
Font.Size = wpara.FontSize(0)

AnzTabs = 7
TabsPerRow = 7
For i% = 1 To 6
    Load picTab(i%)
Next i%

With picTab(0)
    .Height = .TextHeight("Äg") * 2 - 60
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    If (TabsPerRow < AnzTabs) Then
        .Top = .Top + .Height
    End If
End With
With picStammdatenBack(0)
'    .Top = picTab(0).Top + .TextHeight("Äg") * 2 - 60 '(300 * wpara.BildFaktor)
    .Top = picTab(0).Top + picTab(0).Height  ' .TextHeight("Äg") * 2 - 60 '(300 * wpara.BildFaktor)
End With
For i = 0 To 6
    TabNamen(i + 1) = Trim(Mid$(tabOptionen.TabCaption(i), 5))
    TabEnabled(i + 1) = True
    Call PaintTab(i)
Next i
Breite1% = picTab(TabsPerRow - 1).Left + picTab(TabsPerRow - 1).Width + 150

For i% = 0 To 0
    With picStammdatenBack(0)
        .Left = picTab(0).Left
        .Top = picTab(0).Top + (.TextHeight("Äg") * 2 - 60) '(300 * wpara.BildFaktor)
        .Width = Breite1%
        .Height = 3000  'tabOptionen.Height - (.Top - tabOptionen.Top) - 210
        .BorderStyle = 0
    End With
Next i
For i% = 1 To 6
    With picStammdatenBack(i)
        .Left = picStammdatenBack(0).Left
        .Top = picStammdatenBack(0).Top
'        .Width = picStammdatenBack(0).Width
'        .Height = picStammdatenBack(0).Height
        .Width = 900
        .Height = 900
        .BorderStyle = picStammdatenBack(0).BorderStyle
    End With
Next i

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is Label) Then
        c.BackStyle = 0 'duchsichtig
    End If
    
    If (c.Name = picStammdatenBack(0).Name) Then
    ElseIf (c.Container.Name = picStammdatenBack(0).Name) Then
        If (Left(c.Name, 10) = "lblSection") Then
            c.Font.Size = c.Font.Size + 4
        End If
        
        If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
            c.BackColor = vbWhite
        Else
            c.BackColor = RGB(232, 217, 172)
        End If
        
'        If (TypeOf c Is TextBox) Or (TypeOf c Is Label) Or (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Or (TypeOf c Is ComboBox) Then
        If (TypeOf c Is TextBox) Or (TypeOf c Is Label) Or (TypeOf c Is ComboBox) Then
            Font.Name = c.Font.Name
            Font.Size = c.Font.Size
            Font.Bold = c.Font.Bold
            c.Height = TextHeight("Äg") + 60
            If (TypeOf c Is TextBox) Then
            ElseIf (TypeOf c Is Label) Then
                c.Width = TextWidth(c.Caption) + 90
            Else
                c.Width = TextWidth(c.Caption) + 600
            End If
'        ElseIf (TypeOf c Is CheckBox) Then
'            c.Height = 0
'            c.Width = c.Height
'            If (c.Name = "chkStammdaten2") Then
'                With lblChkCaption2(c.Index)
'                    .Caption = c.Caption
'                    .Left = c.Left + 300
'                    .Top = c.Top
'                    .TabIndex = c.TabIndex
'                End With
'            End If
'        ElseIf (TypeOf c Is OptionButton) Then
        ElseIf (TypeOf c Is MSFlexGrid) Then
            With c
'                .FillStyle = flexFillRepeat
'                .row = 0
'                .col = 0
'                .RowSel = .Rows - 1
'                .ColSel = .Cols - 1
'                .CellFontSize = i%
'                .FillStyle = flexFillSingle
'                .row = .FixedRows
'                .col = .FixedCols
                
                .ScrollBars = flexScrollBarNone
                .BorderStyle = 0
                .GridLines = flexGridFlat
                .GridLinesFixed = .GridLines
                .GridColorFixed = .GridColor
                .BackColor = vbWhite
                .BackColorBkg = vbWhite
                .BackColorFixed = RGB(199, 176, 123)
                If (.SelectionMode = flexSelectionFree) Then
                    .BackColorSel = RGB(135, 61, 52)
                    .ForeColorSel = vbWhite '.ForeColor
                Else
                    .BackColorSel = RGB(232, 217, 172)
                    .ForeColorSel = .ForeColor
                End If
                .Appearance = 0
            End With
        End If
    End If
Next
On Error GoTo DefErr




'txtStammdaten(2).Width = TextWidth(String(7, "X")) + 90
'txtStammdaten2(0).Width = TextWidth(String(8, "9")) + 90
'txtStammdaten7(0).Width = txtStammdaten2(0).Width
'txtStammdaten7(12).Width = TextWidth(String(4, "X")) + 90

LblHe% = chkZA_Gruppe(0).Height
TxtHe% = txtZA_Freitext(0).Height
ydiff% = (TxtHe% - LblHe) / Screen.TwipsPerPixelY
ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
'yDiff% = TxtHe% - lblStammdaten(0).Height
        
'xadd = 120
'yAdd = yDiff + 30 + 15

Font.Bold = False   ' True

'''''''''''''''''
'txtOptionen0(1).text = String(38, "A")
'txtOptionen0(2).text = String(38, "A")

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 0

With lblPreise(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With txtPreise(0)
    .Left = lblPreise(0).Left + lblPreise(7).Width + 750
    .Top = lblPreise(0).Top - ydiff
    .Width = TextWidth(String(15, "X"))
End With

For i% = 1 To 9
    With lblPreise(i)
        .Left = lblPreise(0).Left
        .Top = lblPreise(i% - 1).Top + lblPreise(i - 1).Height + 300
        If (i = 2) Or (i = 4) Or (i = 7) Then
            .Top = .Top + 300
        End If
    End With
    If (i <= 8) Then
        With txtPreise(i)
            .Left = txtPreise(0).Left
            .Top = lblPreise(i).Top - ydiff
            .Width = IIf(i = 8, TextWidth(String(35, "X")), txtPreise(0).Width)
        End With
    Else
        With cboVertragsKz
            .Left = txtPreise(0).Left
            .Top = lblPreise(i).Top - ydiff
            .Width = TextWidth(String(35, "X"))
        End With
    End If
Next i%

i = 10
With lblPreise(i)
    .Left = txtPreise(0).Left + txtPreise(0).Width + 1500
    .Top = lblPreise(4).Top
End With
With txtPreise(i)
    .Left = lblPreise(i).Left + lblPreise(i).Width + 600
    .Top = lblPreise(i).Top - ydiff
    .Width = txtPreise(0).Width
End With

i = 11
With lblPreise(i)
    .Left = lblPreise(10).Left
    .Top = lblPreise(5).Top
End With
With txtPreise(i)
    .Left = txtPreise(10).Left
    .Top = lblPreise(i).Top - ydiff
    .Width = txtPreise(0).Width
End With

i = 12
With lblPreise(i)
    .Left = lblPreise(0).Left
    .Top = lblPreise(9).Top + lblPreise(9).Height + 300
End With
With txtPreise(i)
    .Left = txtPreise(0).Left
    .Top = lblPreise(i).Top - ydiff
    .Width = TextWidth(String(30, "X"))
End With

With lblLieferengpass
    .Left = lblPreise(10).Left
    .Top = lblPreise(1).Top
    .Visible = False
End With

With cboVertragsKz
    .Clear
    .AddItem ("ohne ErsatzVO (0)")
    .AddItem ("ASV (1)")
    .AddItem ("Entlassmanagement (4)")
    .AddItem ("Terminservicestellen (7)")
    .AddItem ("nur ErsatzVO (10)")
    .AddItem ("ASV mit ErsatzVO (11)")
    .AddItem ("Entlassmanagement mit ErsatzVO (14)")
    .AddItem ("Terminservicestellen mit ErsatzVO (17)")
End With

With eRezept
    txtPreise(0).text = FormatOrEmpty(xVal(.sGesamtZuzahlung))
    txtPreise(1).text = FormatOrEmpty(xVal(.sGesamtBrutto))
    
    txtPreise(2).text = FormatOrEmpty(xVal(.sBruttoPreis))
    txtPreise(3).text = FormatOrEmpty(xVal(.sMwst))
    
    txtPreise(4).text = FormatOrEmpty(xVal(.sZuzahlung))
    txtPreise(5).text = FormatOrEmpty(xVal(.sMehrkosten))
    txtPreise(6).text = FormatOrEmpty(xVal(.sEigenanteil))
    
    txtPreise(7).text = Format(IIf(Year(.AbgabeDatum) > 2000, .AbgabeDatum, Now), "DDMMYY")
    txtPreise(8).text = .Dosieranweisung
    
    txtPreise(10).text = FormatOrEmpty(xVal(.sBotendienst))
    txtPreise(11).text = FormatOrEmpty(xVal(.sBeschaffungsKosten))
    
    Dim sChargenNr As String
    Dim ChargenNrInd(2) As Integer
    ChargenNrInd(0) = 12
    ChargenNrInd(1) = 14
    ChargenNrInd(2) = 15
    sChargenNr = .ChargenNr + "@"
    For i = 1 To 3
        ind = InStr(sChargenNr, "@")
        If (ind > 0) Then
            txtPreise(ChargenNrInd(i - 1)).text = Left(sChargenNr, ind - 1)
        Else
            Exit For
        End If
        sChargenNr = Mid(sChargenNr, ind + 1)
    Next i
'    txtPreise(12).text = .ChargenNr
    
    If (.ZusatzAttribut(ZusatzattributGruppe_OP_AbgabeImNotdienst).Schluessel = 0) Then
        txtPreise(13).text = FormatOrEmpty(2.5)
    End If
    
    With cboVertragsKz
        For j = 1 To .ListCount
            s = Trim(.List(j - 1))
            If (s <> "") Then
                ind = InStr(s, "(")
                If (ind > 0) Then
                    s = Mid(s, ind + 1)
                    ind = InStr(s, ")")
                    If (ind > 0) Then
                        s = Strings.Left(s, ind - 1)
                        iVal = Val(s)
                        If (iVal = eRezept.Rechtsgrundlage) Then
                            .ListIndex = j - 1
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    End With
'    MsgBox (CStr(.Rechtsgrundlage))
End With

Dim iLblInd(15) As Integer
iLblInd(0) = 2
iLblInd(1) = 3
iLblInd(2) = 10
iLblInd(3) = 11
iLblInd(4) = 13
iLblInd(5) = 1

iLblInd(6) = 7
iLblInd(7) = 8
iLblInd(8) = 9
iLblInd(9) = 12

iLblInd(10) = 4
iLblInd(11) = -1
iLblInd(12) = 5
iLblInd(13) = 6
iLblInd(14) = -1
iLblInd(15) = 0
'ilblind(10)=
'ilblind(11)=
'ilblind()=
            
With lblPreise(iLblInd(0))
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With txtPreise(iLblInd(0))
    .Left = lblPreise(0).Left + lblPreise(7).Width + 750
    .Top = lblPreise(0).Top - ydiff
    .Width = TextWidth(String(15, "X"))
    x2 = .Left + .Width + 1500
End With

For i% = 1 To 9
    With lblPreise(iLblInd(i))
        .Left = lblPreise(iLblInd(0)).Left
        .Top = lblPreise(iLblInd(i% - 1)).Top + lblPreise(iLblInd(i - 1)).Height + 210
        If (i = 2) Or (i = 5) Then
            .Top = .Top + 210
        ElseIf (i = 6) Then
            .Top = .Top + 750
        End If
    End With
    If (i = 8) Then
        With cboVertragsKz
            .Left = txtPreise(iLblInd(0)).Left
            .Top = lblPreise(iLblInd(i)).Top - ydiff
            .Width = TextWidth(String(35, "X"))
        End With
    Else
        With txtPreise(iLblInd(i))
            .Left = txtPreise(iLblInd(0)).Left
            .Top = lblPreise(iLblInd(i)).Top - ydiff
            .Width = IIf(i = 8, TextWidth(String(35, "X")), txtPreise(iLblInd(0)).Width)
        End With
    End If
Next i%

With lblLieferengpass
    .Left = lblPreise(iLblInd(0)).Left
    .Top = txtPreise(iLblInd(5)).Top + txtPreise(iLblInd(5)).Height + 60
    .Visible = False
End With

For i = 14 To 15
    With txtPreise(i)
        .Left = txtPreise(12).Left + (i - 13) * (txtPreise(12).Width + 150)
        .Top = txtPreise(12).Top
        .Width = txtPreise(12).Width
        .Visible = (eRezept.Anzahl > (i - 13))
    End With
Next i


For i% = 10 To 15
    If (iLblInd(i) >= 0) Then
        With lblPreise(iLblInd(i))
            .Left = x2
            .Top = lblPreise(iLblInd(i% - 10)).Top
        End With
        With txtPreise(iLblInd(i))
            .Left = x2 + lblPreise(iLblInd(10)).Width + 750
            .Top = lblPreise(iLblInd(i)).Top - ydiff
            .Width = txtPreise(iLblInd(0)).Width
        End With
    End If
Next i%

'i = 10
'With lblPreise(i)
'    .Left = txtPreise(0).Left + txtPreise(0).Width + 1500
'    .Top = lblPreise(4).Top
'End With
'With txtPreise(i)
'    .Left = lblPreise(i).Left + lblPreise(i).Width + 600
'    .Top = lblPreise(i).Top - ydiff
'    .Width = txtPreise(0).Width
'End With
'
'i = 11
'With lblPreise(i)
'    .Left = lblPreise(10).Left
'    .Top = lblPreise(5).Top
'End With
'With txtPreise(i)
'    .Left = txtPreise(10).Left
'    .Top = lblPreise(i).Top - ydiff
'    .Width = txtPreise(0).Width
'End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 1

With lblVO(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With txtVO(0)
    .Left = lblVO(0).Left + lblVO(4).Width + 750
    .Top = lblVO(0).Top - ydiff
    .Width = TextWidth(String(40, "X"))
End With
With txtAbgabe(0)
    .Left = txtVO(0).Left + txtVO(0).Width + 750
    .Top = lblVO(0).Top - ydiff
    .Width = TextWidth(String(40, "X"))
End With

For i% = 1 To 7
    With lblVO(i)
        .Left = lblVO(0).Left
        .Top = lblVO(i% - 1).Top + lblVO(i - 1).Height + 300
        If (i = 1) Or (i = 7) Then
            .Top = .Top + 300
        End If
    End With
    With txtVO(i)
        .Left = txtVO(0).Left
        .Top = lblVO(i).Top - ydiff
        .Width = txtVO(0).Width
    End With
    With txtAbgabe(i)
        .Left = txtAbgabe(0).Left
        .Top = lblVO(i).Top - ydiff
        .Width = txtAbgabe(0).Width
    End With
Next i%
With cboVerordnungstyp
    .Left = txtAbgabe(0).Left
    .Top = txtAbgabe(0).Top
    .Width = txtAbgabe(0).Width
End With

With cboVerordnungstyp
    .Clear
    .AddItem ("Unbekannt   (0)")
    .AddItem ("PZNVerordnung   (1)")
    .AddItem ("Wirkstoffverordnung   (2)")
    .AddItem ("Rezepturverordnung   (3)")
    .AddItem ("Freitextverordnung   (4)")
End With

With eRezept
    With cboVerordnungstyp
        For j = 1 To .ListCount
            s = Trim(.List(j - 1))
            If (s <> "") Then
                ind = InStr(s, "(")
                If (ind > 0) Then
                    s = Mid(s, ind + 1)
                    ind = InStr(s, ")")
                    If (ind > 0) Then
                        s = Strings.Left(s, ind - 1)
                        iVal = Val(s)
                        If (iVal = eRezept.Verordnung.VerordnungsTyp) Then
                            txtVO(0).text = Trim(.List(j - 1))
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    End With

    txtVO(1).text = .Verordnung.pzn
    txtVO(2).text = .Verordnung.Verordnungstext
    txtVO(3).text = .Verordnung.Darreichungsform
    txtVO(4).text = .Verordnung.sPackungsgroesse
    txtVO(5).text = .Verordnung.einheit
    txtVO(6).text = .Verordnung.Normgroesse
    
    txtVO(7).text = .AnzahlVerordnet
    
'    chkVO_Impfstoff.Checked = .Verordnung.Impfstoff

    With cboVerordnungstyp
        For j = 1 To .ListCount
            s = Trim(.List(j - 1))
            If (s <> "") Then
                ind = InStr(s, "(")
                If (ind > 0) Then
                    s = Mid(s, ind + 1)
                    ind = InStr(s, ")")
                    If (ind > 0) Then
                        s = Strings.Left(s, ind - 1)
                        iVal = Val(s)
                        If (iVal = eRezept.Abgabe.VerordnungsTyp) Then
                            .ListIndex = j - 1
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    End With

    'txtABG_Verordnungstyp.Text = .Abgabe.VerordnungsTyp.ToString
    txtAbgabe(1).text = .Abgabe.pzn
    txtAbgabe(2).text = .Abgabe.Verordnungstext
    txtAbgabe(3).text = .Abgabe.Darreichungsform
    txtAbgabe(4).text = .Abgabe.sPackungsgroesse
    txtAbgabe(5).text = .Abgabe.einheit
    txtAbgabe(6).text = .Abgabe.Normgroesse
    txtAbgabe(7).text = .Anzahl
'    chkABG_Impfstoff.Checked = .Abgabe.Impfstoff
End With

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 2

With chkZA_Gruppe(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With cboZA_Schluessel(0)
    .Left = chkZA_Gruppe(0).Left + chkZA_Gruppe(15).Width + 600
    .Top = chkZA_Gruppe(0).Top - ydiff
    .Width = TextWidth(String(30, "X"))
End With
With txtZA_Freitext(0)
    .Left = cboZA_Schluessel(0).Left + cboZA_Schluessel(0).Width + 600
    .Top = chkZA_Gruppe(0).Top - ydiff
    .Width = TextWidth(String(30, "X"))
End With

For i% = 1 To 15
    With chkZA_Gruppe(i)
        .Left = chkZA_Gruppe(0).Left
        .Top = chkZA_Gruppe(i% - 1).Top + chkZA_Gruppe(i% - 1).Height + 150 ' + 300
    End With
    With cboZA_Schluessel(i)
        .Left = cboZA_Schluessel(i - 1).Left
        .Top = chkZA_Gruppe(i).Top - ydiff
        .Width = cboZA_Schluessel(0).Width
    End With
    With txtZA_Freitext(i%)
        .Left = txtZA_Freitext(0).Left
        .Top = chkZA_Gruppe(i).Top - ydiff
        .Width = txtZA_Freitext(0).Width
    End With
Next i%
With txtZA_Datum(10)
    .Left = txtZA_Freitext(0).Left + txtZA_Freitext(0).Width + 750
    .Top = txtZA_Freitext(10).Top
    .Width = TextWidth(String(8, "X"))
End With
With txtZA_Uhrzeit(10)
    .Left = txtZA_Datum(10).Left + txtZA_Datum(10).Width + 150
    .Top = txtZA_Datum(10).Top
    .Width = TextWidth(String(5, "X"))
End With
With txtZA_Datum(12)
    .Left = txtZA_Datum(10).Left
    .Top = txtZA_Freitext(12).Top
    .Width = TextWidth(String(8, "X"))
End With
With lblSpenderPzn
    .Left = txtZA_Freitext(15).Left
    .Top = txtZA_Freitext(15).Top + txtZA_Freitext(15).Height
    .Width = txtZA_Freitext(15).Width * 2
End With

With cboZA_Schluessel(0)
    .Clear
    .AddItem ("nicht betroffen (0)")
    .AddItem ("Generika (1)")
    .AddItem ("Solitär (2)")
    .AddItem ("Mehrfach-VO (3)")
    .AddItem ("autIdem gesetzt (4)")
    .AddItem ("Produkt der Substitutionsausschlussliste (5)")
End With
For i = 1 To 3
    With cboZA_Schluessel(i)
        .Clear
        .AddItem ("nicht relevant (0)")
        .AddItem ("ja, abgegeben(1)")
        .AddItem ("nein, Nicht-Verfügbarkeit + FT (2)")
        .AddItem ("nein, dringender Fall  (3)")
        .AddItem ("nein, sonstige Bedenken + FT  (4)")
        If (i = 3) Then
            .AddItem ("nicht abgegeben (5)")
        End If
    End With
Next
With cboZA_Schluessel(4)
    .Clear
    .AddItem ("nach §129 Abs.4c SGB V Generika (1)")
    .AddItem ("nach Rabattvertrag (2)")
End With
With cboZA_Schluessel(14)
    .Clear
    .AddItem ("gebührenpflichtig (0)")
    .AddItem ("gebührenfrei (1)")
End With

For i = 1 To 16
    txtZA_Freitext(i - 1).MaxLength = 100
Next i

With eRezept
    For i = 1 To 16
        chkZA_Gruppe(i - 1).Value = Abs(.ZusatzAttribut(i).Schluessel >= 0)
        If (i <= 5) Or (i = 15) Then
'            Dim cbo As NlComboBox = rgbZusatzAttribute.Controls("cboZA_Schluessel_" + i.ToString)
            With cboZA_Schluessel(i - 1)
                For j = 1 To .ListCount
                    s = Trim(.List(j - 1))
                    If (s <> "") Then
                        ind = InStr(s, "(")
                        If (ind > 0) Then
                            s = Mid(s, ind + 1)
                            ind = InStr(s, ")")
                            If (ind > 0) Then
                                s = Strings.Left(s, ind - 1)
                                iVal = Val(s)
                                If (iVal = eRezept.ZusatzAttribut(i).Schluessel) Then
                                    .ListIndex = j - 1
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End With
        End If
        If (Trim(.ZusatzAttribut(i).Freitext) <> "") Then
            txtZA_Freitext(i - 1).text = .ZusatzAttribut(i).Freitext
        End If
        If (Year(.ZusatzAttribut(i).Datum) >= 2000) Then
            txtZA_Datum(i - 1).text = Format(.ZusatzAttribut(i).Datum, "DDMMYY")
            txtZA_Uhrzeit(i - 1).text = Format(.ZusatzAttribut(i).Datum, "HHmm")
        End If
    Next
End With

Call CheckLieferengpass
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 3

With chkRAE_Art(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With cbokRAE_RueckspracheArzt(0)
    .Left = chkRAE_Art(0).Left + chkRAE_Art(2).Width + 750
    .Top = chkRAE_Art(0).Top - ydiff
    .Width = TextWidth(String(30, "X"))
End With
With txtRAE_Datum(0)
    .Left = cbokRAE_RueckspracheArzt(0).Left + cbokRAE_RueckspracheArzt(0).Width + 750
    .Top = chkRAE_Art(0).Top - ydiff
    .Width = TextWidth(String(8, "X"))
End With
With txtRAE_Uhrzeit(0)
    .Left = txtRAE_Datum(0).Left + txtRAE_Datum(0).Width + 150
    .Top = txtRAE_Datum(0).Top
    .Width = TextWidth(String(5, "X"))
End With
With txtRAE_Dokumentation(0)
    .Left = txtRAE_Uhrzeit(0).Left + txtRAE_Uhrzeit(0).Width + 750
    .Top = txtRAE_Datum(0).Top
    .Width = TextWidth(String(30, "X"))
End With

For i% = 1 To 11
    With chkRAE_Art(i)
        .Left = chkRAE_Art(0).Left
        .Top = chkRAE_Art(i% - 1).Top + chkRAE_Art(i% - 1).Height + 300
    End With
    With cbokRAE_RueckspracheArzt(i)
        .Left = cbokRAE_RueckspracheArzt(i - 1).Left
        .Top = chkRAE_Art(i).Top - ydiff
        .Width = cbokRAE_RueckspracheArzt(0).Width
    End With
    With txtRAE_Datum(i%)
        .Left = txtRAE_Datum(0).Left
        .Top = chkRAE_Art(i).Top - ydiff
        .Width = txtRAE_Datum(0).Width
    End With
    With txtRAE_Uhrzeit(i%)
        .Left = txtRAE_Uhrzeit(0).Left
        .Top = chkRAE_Art(i).Top - ydiff
        .Width = txtRAE_Uhrzeit(0).Width
    End With
    With txtRAE_Dokumentation(i%)
        .Left = txtRAE_Dokumentation(0).Left
        .Top = chkRAE_Art(i).Top - ydiff
        .Width = txtRAE_Dokumentation(0).Width
    End With
Next i%

For i = 0 To 11
    With cbokRAE_RueckspracheArzt(i)
        .Clear
        .AddItem ("RueckspracheErfolgt (0)")
        .AddItem ("DringenderFallRueckspracheWarNichtMoeglich (1)")
        .AddItem ("NichtErforderlich (2)")
    End With
Next

With eRezept
    For i = 1 To 12
        chkRAE_Art(i - 1).Value = Abs(.RezeptAenderung(i).RueckspracheArzt >= 0)
            
        With cbokRAE_RueckspracheArzt(i - 1)
            For j = 1 To .ListCount
                s = Trim(.List(j - 1))
            
                If (s <> "") Then
                    ind = InStr(s, "(")
                    If (ind > 0) Then
                        s = Mid(s, ind + 1)
                        ind = InStr(s, ")")
                        If (ind > 0) Then
                            s = Strings.Left(s, ind - 1)
                            iVal = Val(s)
                            If (iVal = eRezept.RezeptAenderung(i).RueckspracheArzt) Then
                                .ListIndex = j - 1
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        End With
        
        If (Trim(.RezeptAenderung(i).Dokumentation) <> "") Then
            txtRAE_Dokumentation(i - 1).text = Trim(.RezeptAenderung(i).Dokumentation)
        End If
        If (Year(.RezeptAenderung(i).DatumAenderung) >= 2000) Then
            txtRAE_Datum(i - 1).text = Format(.RezeptAenderung(i).DatumAenderung, "DDMMYY")
            txtRAE_Uhrzeit(i - 1).text = Format(.RezeptAenderung(i).DatumAenderung, "HHmm")
        End If
    Next
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 4

With wbVerordnung
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    .Width = txtRAE_Dokumentation(11).Left + txtRAE_Dokumentation(11).Width - .Left
    .Height = txtRAE_Dokumentation(11).Top + txtRAE_Dokumentation(11).Height - .Top

    Dim sDatei$, sBundleKBV$, sXslt$
    Dim l&
    
    If (eRezept.Mehrfach) Then
        With eRezept.MehrfachVerordnung
            MsgBox (.sZaehler + " / " + .sNenner + " " + CStr(.Start) + " " + CStr(.Ende))
        End With
    End If
'    MsgBox (eRezept.Kostentraeger.Typ)
    
    sBundleKBV = eRezept.sBundleKBV
    sDatei = CurDir() + "\eVerordnung.xml"
    l = writeOut(sBundleKBV, sDatei)
    
'    l = Transformieren(sDatei, CurDir() + "\ERP_Stylesheet.xslt", CurDir() + "\eVerordnung.htm")
    
    Dim sVersionKBV As String
    Dim sPriv As String
    sVersionKBV = eRezept.Version_FhirERezept
    sVersionKBV = Replace(sVersionKBV, ".", "_")
    
    sPriv = IIf(eRezept.RezeptTypId = 200, "p", "")
    If (eRezept.Kostentraeger.Typ = "SEL") Then
        sPriv = "g"
    End If
    
    sXslt = CurDir + "\erezepte\xslt\ERP_Stylesheet." + sVersionKBV + sPriv + ".xslt"
    If (FileExist(sXslt)) Then
'        MsgBox (sXslt)
    Else
        Call MsgBox("Problem:" + vbCrLf + vbCrLf + "Datei '" + sXslt + "' nicht vorhanden!", vbInformation)
        sXslt = CurDir() + "\ERP_Stylesheet.xslt"
    End If
    lblVerordnung.Caption = "Xslt: " + sXslt
    l = Transformieren(sDatei, sXslt, CurDir() + "\eVerordnung.htm")
    
    With wbVerordnung
        .Visible = False
'    '    .Navigate ("about:blank")
'        .Width = frmAction.Width '/ 2
'        .Height = frmAction.Height '/ 2
'        .Width = 950 * 15
'        .Height = 550 * 15
'
        .Navigate (CurDir() + "\eVerordnung.htm")
        
        .Visible = True
    End With
    With lblVerordnung
        .Left = wbVerordnung.Left
        .Top = wbVerordnung.Top + wbVerordnung.Height + 90
        .Width = wbVerordnung.Width
    End With
      
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 5

With wbAbgabedaten
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    .Width = txtRAE_Dokumentation(11).Left + txtRAE_Dokumentation(11).Width - .Left
    .Height = txtRAE_Dokumentation(11).Top + txtRAE_Dokumentation(11).Height - .Top

    Dim sVersionAbgabedaten As String
    Dim sBundleDAV$
    
'    h = InputBox("Anzahl verordnete Packungen: ", "eRezept", eRezept.Anzahl)
'    If (h <> "") Then
'        eRezept.Anzahl = h
'    End If
   
    sBundleDAV = FD_OP.TI_RezeptAbgabedaten(eRezept, sVersionAbgabedaten)
    'MsgBox (sVersionAbgabedaten)
    sVersionAbgabedaten = Replace(sVersionAbgabedaten, ".", "_")
    'MsgBox (sVersionAbgabedaten)
    
    sDatei = CurDir() + "\eAbgabe.xml"
    l = writeOut(sBundleDAV, sDatei)

'    l = Transformieren(sDatei, CurDir() + "\eAbgabedaten.xslt", CurDir() + "\eAbgabe.htm")
    
    sXslt = CurDir + "\erezepte\xslt\eAbgabedatensatz." + sVersionAbgabedaten + sPriv + ".xslt"
    If (FileExist(sXslt)) Then
'        MsgBox (sXslt)
    Else
        Call MsgBox("Problem:" + vbCrLf + vbCrLf + "Datei '" + sXslt + "' nicht vorhanden!", vbInformation)
        sXslt = CurDir() + "\eAbgabedaten.xslt"
    End If
    lblAbgabedaten.Caption = "Xslt: " + sXslt
    l = Transformieren(sDatei, sXslt, CurDir() + "\eAbgabe.htm")
    
    With wbAbgabedaten
        .Visible = False
'        .Width = 900 * 15
'        .Height = 550 * 15
'        .Left = wbVerordnung.Left + wbVerordnung.Width + 60
'        .Top = wbVerordnung.Top
        .Navigate (CurDir() + "\eAbgabe.htm")
        .Visible = True
    End With
    With lblAbgabedaten
        .Left = wbAbgabedaten.Left
        .Top = wbAbgabedaten.Top + wbAbgabedaten.Height + 90
        .Width = wbAbgabedaten.Width
    End With
    
'    Call FD_OP.PKV_Ausdruck(eRezept, True)
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 6

With lblXML(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With txtXML(0)
    .Left = lblXML(0).Left + lblXML(2).Width + 750
    .Top = lblXML(0).Top - ydiff
    .Width = TextWidth(String(40, "X"))
End With

For i% = 1 To 5
    With lblXML(i)
        .Left = lblXML(0).Left
        .Top = lblXML(i% - 1).Top + lblXML(i - 1).Height + 300
        If (i >= 3) Then
            .Top = .Top + 1200
        End If
    End With
    With txtXML(i)
        .Left = txtXML(0).Left
        .Top = lblXML(i).Top - ydiff
        .Width = IIf(i >= 3, TextWidth(String(80, "X")), txtXML(0).Width)
        If (i >= 3) Then
            .Height = .Height + 900
        End If
    End With
Next i%
For i% = 6 To 7
    With lblXML(i)
        .Left = txtXML(3).Left + txtXML(3).Width + 600
        .Top = lblXML(i% - 2).Top
        .Caption = ""
    End With
Next

Dim sSQL As String
Dim Rec2 As New ADODB.Recordset
sSQL = "Select * FROM TI_eRezepte"
sSQL = sSQL + " WHERE TI_eRezepte.TaskId='" + eRezept.TaskId + "'"
FabsErrf = VerkaufAdoDB.OpenRecordset(Rec2, sSQL, 0)
If (FabsErrf = 0) Then
'    sQuittungGEM = sCheckDBNull(dr.Item("QuittungXML"))

    txtXML(0).text = CheckNullStr(Rec2!TaskId)
    txtXML(1).text = CheckNullStr(Rec2!AccessCode)
    txtXML(2).text = CheckNullStr(Rec2!Secret)
    
    txtXML(3).text = CheckNullStr(Rec2!Bundle)
    
    h = CheckNullStr(Rec2!DispenseXML)
    
    ind = InStr(UCase(h), "<?XML V")
    If (ind > 0) Then
        h = Mid(h, ind)
    Else
        ind = InStr(UCase(h), "<MEDICA")
        If (ind > 0) Then
            h = Mid(h, ind)
        End If
    End If
    txtXML(4).text = h
    
    txtXML(5).text = CheckNullStr(Rec2!eAbgabe)
    
    lblXML(6).Caption = "(Quittung Len: " + CStr(Len(CheckNullStr(Rec2!QuittungXML))) + ")"
    lblXML(7).Caption = "(eDispensierung Len: " + CStr(Len(CheckNullStr(Rec2!eDispensierung))) + ")"
End If
Rec2.Close
            


'txtOptionen0(0).text = OrgRezApoNr$
'OrgRezApoDruckName$ = RezApoDruckName$
'txtOptionen0(1).text = Trim(Left$(RezApoDruckName$ + Space$(38), 38))
'OrgBtmRezDruckName$ = BtmRezDruckName$
'txtOptionen0(2).text = Trim(Left$(BtmRezDruckName$ + Space$(50), 50))
'
'txtOptionen0(3).text = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")
'
'chkOptionen0(0).Value = Abs(RezepturMitFaktor%)
'chkOptionen0(1).Value = Abs(BtmAlsZeile%)
'chkOptionen0(2).Value = Abs(RezepturDruck%)
'chkOptionen0(3).Value = Abs(AvpTeilnahme%)

'''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 1
'Call ActProgram.FlxOptionenBefuellen(flxOptionen1(0))
'With flxOptionen1(0)
'    Breite1% = 0
'    For i% = 0 To (.Rows - 1)
'        Breite2% = TextWidth(.TextMatrix(i%, 0))
'        If (Breite2% > Breite1%) Then Breite1% = Breite2%
'    Next i%
'    .ColWidth(0) = Breite1% + 150
'    .ColWidth(1) = TextWidth("00000")
'    .ColWidth(2) = wpara.FrmScrollHeight
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) '+ 90
'    .Height = .RowHeight(0) * 11 '+ 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .ScrollBars = flexScrollBarVertical
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 2
'With flxOptionen1(1)
'    .Cols = 3
'    .Rows = 2
'    .FixedRows = 1
'    .Rows = 1
'
'    .FormatString = "^Tätigkeit|^Personal|^ "
'
'    .ColWidth(0) = TextWidth(String(20, "A")) + 150
'    .ColWidth(1) = TextWidth(String(30, "A"))
'    .ColWidth(2) = 0
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) '+ 90
'    .Height = .RowHeight(0) * 11 '+ 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'
'    For j = 0 To 1
'        If (j = 0) Then
'            h$ = "Rezeptspeicher"
'        Else
'            h$ = "Importkontrolle"
'        End If
'
'        For i% = 0 To (AnzTaetigkeiten% - 1)
'            h2$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
'            If (UCase(h) = UCase(h2)) Then
'                h2$ = ""
'                h3 = ""
'                For k% = 0 To 79
'                    If (Taetigkeiten(i%).pers(k%) > 0) Then
'                        h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
'                    Else
'                        Exit For
'                    End If
'                Next k%
'                If (k% = 1) Then
'                    ind% = Taetigkeiten(i%).pers(0)
'                    h3$ = RTrim$(para.Personal(ind%))
'                ElseIf (k% > 1) Then
'                    h3$ = "mehrere (" + Mid$(Str$(k%), 2) + ")"
'                End If
'                h3$ = h3$ + vbTab + h2$
'            End If
'        Next i%
'
'        .AddItem h$ + vbTab + h3
'    Next j
'    .row = 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 3
'With flxOptionen1(2)
'    .Rows = 13
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 1
'    .TextMatrix(0, 0) = "Monat"
'    .TextMatrix(0, 1) = "Datum"
'    .ColWidth(0) = TextWidth(String$(10, "A"))
'    .ColWidth(1) = TextWidth(String$(10, "0"))
'    AbrechDatenRec.MoveFirst
'    For i% = 1 To 12
'        .TextMatrix(i%, 0) = Format(CDate("01." + CStr(i%) + ".2002"), "MMMM")
'        .TextMatrix(i%, 1) = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
'        AbrechDatenRec.MoveNext
'    Next i%
'
'    .Width = .ColWidth(0) + .ColWidth(1) '+ 90
'    .Height = .RowHeight(0) * 13 '+ 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 4
'With flxOptionen1(3)
'    .Rows = 11
'    .Cols = 5
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<SonderPzn|<KK-Bez|<KassenNr|<Status|<GültigBis"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(13, "X"))
'    .ColWidth(2) = TextWidth(String$(10, "9"))
'    .ColWidth(3) = TextWidth(String$(10, "9"))
'    .ColWidth(4) = TextWidth(String$(10, "9"))
''    .ColWidth(5) = wpara.FrmScrollHeight
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% '+ 90
'    .Height = .RowHeight(0) * .Rows '+ 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    For i% = 1 To AnzSonderBelege%
'        .TextMatrix(i%, 0) = SonderBelege(i% - 1).pzn
'        .TextMatrix(i%, 1) = SonderBelege(i% - 1).KkBez
'        .TextMatrix(i%, 2) = SonderBelege(i% - 1).KassenId
'        .TextMatrix(i%, 3) = SonderBelege(i% - 1).Status
'        .TextMatrix(i%, 4) = SonderBelege(i% - 1).GültigBis
'    Next i%
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 5
'With flxOptionen1(4)
'    .Rows = 10
'    .Cols = 4
'    .FixedRows = 1
'    .FixedCols = 2
'
'    .FormatString = "<SonderPzn|<Bezeichnung|>ArbeitsPreis GKV|>Privat"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(42, "X"))
'    .ColWidth(2) = TextWidth(String$(17, "9"))
'    .ColWidth(3) = TextWidth(String$(17, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% '+ 90
'    .Height = .RowHeight(0) * .Rows '+ 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    .Rows = .FixedRows
''    For i% = 0 To UBound(ParenteralPzn)
'    For i% = 0 To 7
'        h = ParenteralPzn(i) + vbTab + ParenteralTxt(i) + vbTab + Format(ParenteralPreis(i), "0.00") + vbTab + Format(ParenteralPreis(i + 8), "0.00")
'        .AddItem h$
'    Next i%
'
'    .row = .FixedRows
'    .col = .Cols - 1
'    .ColSel = .col
'End With
'
'With lblOptionen5(0)
'    .Left = wpara.LinksX
'    .Top = flxOptionen1(4).Top + flxOptionen1(4).Height + 450
'    txtOptionen5(0).Left = .Left + .Width + 150
'    txtOptionen5(0).Top = .Top - ydiff
'End With
'For i% = 1 To 1
'    With lblOptionen5(i%)
'        .Left = lblOptionen5(0).Left
'        .Top = lblOptionen5(i% - 1).Top + lblOptionen0(i% - 1).Height + 150
'    End With
'    With txtOptionen5(i%)
'        .Left = txtOptionen5(0).Left
'        .Top = lblOptionen5(i%).Top - ydiff
'    End With
'Next i%
'For i = 0 To 1
'    txtOptionen5(i).text = Format(ParEnteralAufschlag(i), "0.00")
'Next i
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 6
'With flxOptionen1(5)
'    .Rows = 11
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<AbgabeArt|>Preis"
'
'    .ColWidth(0) = TextWidth(String$(30, "X"))
'    .ColWidth(1) = TextWidth(String$(10, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% '+ 90
'    .Height = .RowHeight(0) * .Rows '+ 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    Call ActProgram.LadeOptionenAbgabeKosten(flxOptionen1(5))
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
'''''''''''''''''''''''''''''''''''


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Breite1 = txtZA_Freitext(0).Left + txtZA_Freitext(0).Width + 900
Breite1 = txtRAE_Dokumentation(0).Left + txtRAE_Dokumentation(0).Width + 900

Hoehe1% = chkZA_Gruppe(15).Top + chkZA_Gruppe(15).Height
Hoehe2% = 0 ' flxOptionen1(2).Top + flxOptionen1(2).Height
If (Hoehe2% > Hoehe1%) Then
    Hoehe1% = Hoehe2%
End If

With picStammdatenBack(0)
    .Width = Breite1 ' lblOptionenZusatz0(1).Left + lblOptionenZusatz0(1).Width + 2 * wpara.LinksX
    .Height = Hoehe1 + 3 * wpara.TitelY 'chkOptionen0(2).Top + chkOptionen0(2).Height + 2 * wpara.TitelY
            
    .BackColor = RGB(232, 217, 172)
    Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
    picStammdatenBack(0).Line (0, .ScaleHeight - (30 * wpara.BildFaktor) * 15)-(.ScaleWidth, .ScaleHeight), .BackColor, BF

    x = lblPreise(iLblInd(0)).Left
    y = txtPreise(iLblInd(5)).Top - 180
    picStammdatenBack(0).Line (x, y)-(txtPreise(iLblInd(5)).Left + txtPreise(iLblInd(5)).Width, y + 30), vbGrayText, BF
    
'    x = lblPreise(iLblInd(10)).Left
'    picStammdatenBack(0).Line (x, y)-(txtPreise(iLblInd(10)).Left + txtPreise(iLblInd(10)).Width, y + 30), vbGrayText, BF
End With

For i% = 1 To 6
    With picStammdatenBack(i)
        .Left = picStammdatenBack(0).Left
        .Top = picStammdatenBack(0).Top
'        .Width = picStammdatenBack(0).Width
'        .Height = picStammdatenBack(0).Height
        .Width = 900
        .Height = 900
        
        .BorderStyle = picStammdatenBack(0).BorderStyle
        
        BitBlt .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, picStammdatenBack(0).hdc, 0, 0, SRCCOPY
    End With
Next i

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Call DefErrPop
End Sub

Sub PaintTab(Index)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintTab")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, TextHe%
Dim RetVal&, lColor&, bColor&(1)
Dim xStart#, x#, y#
Dim h$
Dim c As Control

'Call DefErrPop: Exit Sub
With picTab(Index)
    .Visible = False
    If (Index = 0) Then
'        .Left = picStammdatenBack(0).Left + 15
    Else
        .Left = picTab(Index - 1).Left + picTab(Index - 1).Width '+ 90
    End If
    .Top = picTab(0).Top
    
    .Width = 3000   ' tabOptionen.Width - 3 * wpara.LinksX
    .Height = picTab(0).Height
    
    If (Index >= TabsPerRow) Then
        .Top = .Top - .Height
        
        If (Index = TabsPerRow) Then
            .Left = picTab(0).Left + 210
        Else
            .Left = picTab(Index - 1).Left + picTab(Index - 1).Width '+ 90
        End If
    End If
    
    .BorderStyle = 0
    
    .Enabled = False
    If (TabEnabled(0) = 0) Then
        bColor(0) = RGB(150, 150, 150)
        bColor(1) = RGB(165, 165, 165)
    ElseIf (Index = AktTab) Then
        bColor(0) = picStammdatenBack(0).BackColor
        bColor(1) = RGB(242, 237, 192)  'bColor(0)
        .Enabled = True
    ElseIf (TabEnabled(Index + 1)) Then
        bColor(0) = RGB(199, 176, 123)
        bColor(1) = RGB(214, 191, 138)
        .Enabled = True
    Else
        bColor(0) = RGB(150, 150, 150)
        bColor(1) = RGB(165, 165, 165)
    End If
    
    .BackColor = RGB(200, 200, 200)
    
    .FillStyle = vbSolid
    .FillColor = bColor(0)
    .ForeColor = bColor(0)
    RoundRect .hdc, 0, 0, .Width, .Height, 10, 10
    
    .FillColor = bColor(1)
    .ForeColor = .FillColor
    RoundRect .hdc, 2, 2, .Width, 10, 10, 10
    picTab(Index).Line (30, 90)-(.Width, .Height / 2), .FillColor, BF
    
    TextHe = .TextHeight("Äg")
    
    .CurrentX = 90
    .CurrentY = (.Height - TextHe) / 2
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = vbWhite
    ElseIf (Index = AktTab) Then
        .ForeColor = RGB(135, 61, 52) ' vbWhite
    Else
        .ForeColor = vbWhite
    End If
    .FillStyle = vbSolid
    .FillColor = .ForeColor
    RoundRect .hdc, .CurrentX / Screen.TwipsPerPixelX, (.CurrentY) / Screen.TwipsPerPixelY, (.CurrentX + TextHe) / Screen.TwipsPerPixelX, (.CurrentY + TextHe) / Screen.TwipsPerPixelY, 5, 5
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = .BackColor
    ElseIf (Index = AktTab) Then
        .ForeColor = vbWhite
    Else
        .ForeColor = .BackColor
    End If
    h$ = CStr(Index + 1)
    .CurrentX = 90 + (TextHe - .TextWidth(h$)) / 2
    picTab(Index).Print h$;
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = vbWhite
    ElseIf (Index = AktTab) Then
        .ForeColor = RGB(135, 61, 52) ' vbWhite
    Else
        .ForeColor = vbWhite
    End If
    h = TabNamen(Index + 1)
    .CurrentX = 60 + TextHe + 150
    picTab(Index).Print h$;
    
    .Width = .CurrentX + 1800

    .ForeColor = RGB(200, 200, 200) ' vbWhite
    
    xStart = .CurrentX + 150
    For i = 0 To 180
        y = -Sin((i + 90) * PI / 180) + 1
        y = y * .ScaleHeight / 2
        x = i * 510# / 180
        picTab(Index).Line (xStart + x, 0)-(xStart + x, y), .ForeColor
        SetPixel .hdc, (xStart + x) / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY + 1, bColor(0)
    Next i
    .Width = xStart + 480   ' 600
    .Visible = True
    
    .Refresh
End With

With picStammdatenBack(Index)
'    If (IstVorlage) Then
'        Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(236, 215, 53))
'    Else
'        Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
'    End If

    If (Index = 0) Then
    ElseIf (Index = AktTab) Then
'        .Width = tabOptionen.Width - 30 '- 3 * wpara.LinksX
'        .Height = tabOptionen.Height - (.Top - tabOptionen.Top) - 210
        .Left = picStammdatenBack(0).Left
        .Width = picStammdatenBack(0).Width
        .Height = picStammdatenBack(0).Height

        .BackColor = RGB(232, 217, 172)
        Call wpara.FillGradient(picStammdatenBack(Index), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
        picStammdatenBack(Index).Line (0, .ScaleHeight - (50 * wpara.BildFaktor) * 15)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
'        .BackColor = picStammdatenBack(0).BackColor
'        BitBlt .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, picStammdatenBack(0).hdc, 0, 0, SRCCOPY
    Else
        .Width = 900
        .Height = 900
    End If
    
    .TabStop = False
    .Visible = (Index = AktTab)
End With


''''''''
If (Index = AktTab) Then
    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (c.Container.Name = picStammdatenBack(0).Name) Then
                If (c.Container.Index = Index) Then
                    If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
        '                If (TypeOf c Is ComboBox) Then
        '                    Call wpara.ControlBorderless(c)
        '                ElseIf (c.Appearance = 1) Then
        '                    Call wpara.ControlBorderless(c, 2, 2)
        '                Else
        '                    Call wpara.ControlBorderless(c, 1, 1)
        '                End If
        
                        If (c.Enabled) And (c.Locked = 0) Then
                            c.BackColor = vbWhite
                        Else
                            c.BackColor = lblPreise(0).BackColor
                        End If
                        
                        If (c.Visible) Then
                            With c.Container
                                .ForeColor = RGB(180, 180, 180) ' vbWhite
                                .FillStyle = vbSolid
                                .FillColor = c.BackColor
                
                                RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                            End With
                        End If
                    End If
                End If
            End If
        End If
    Next
    picStammdatenBack(Index).Refresh
    On Error GoTo DefErr
End If

Call DefErrPop: Exit Sub
'''''''''''''''''''''''''''''''''

picStammdatenBack(Index).Visible = (Index = AktTab)
If (Index = AktTab) Then
    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (c.Container.Name = picStammdatenBack(0).Name) Then
                If (c.Container.Index = Index) Then
                    If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
        '                If (TypeOf c Is ComboBox) Then
        '                    Call wpara.ControlBorderless(c)
        '                ElseIf (c.Appearance = 1) Then
        '                    Call wpara.ControlBorderless(c, 2, 2)
        '                Else
        '                    Call wpara.ControlBorderless(c, 1, 1)
        '                End If
        
                        If (c.Enabled) And (c.Locked = 0) Then
                            c.BackColor = vbWhite
                        Else
                            c.BackColor = lblPreise(0).BackColor
                        End If
                        
                        If (c.Visible) Then
                            With c.Container
                                .ForeColor = RGB(180, 180, 180) ' vbWhite
                                .FillStyle = vbSolid
                                .FillColor = c.BackColor
                
                                RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                            End With
                        End If
                    End If
                End If
            End If
        End If
    Next
    picStammdatenBack(Index).Refresh
    On Error GoTo DefErr
End If

'tabOptionen.Visible = False

Call DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
If (y <= wpara.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
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

If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If

Call DefErrPop
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdQuittungErrneut_Click()
Call cmdQuittungErrneut_Click
End Sub

Private Sub nlcmdDruckVerordnung_Click()
Call cmdDruckVerordnung_Click
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
'    ElseIf (KeyAscii = Asc("<")) And (nlcmdImport(0).Visible) Then
''        Call nlcmdChange_Click(0)
'        nlcmdImport(0).Value = 1
'    ElseIf (KeyAscii = Asc(">")) And (nlcmdImport(1).Visible) Then
''        Call nlcmdChange_Click(1)
'        nlcmdImport(1).Value = 1
    End If
End If
    
If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub

Private Sub txtPreise_Change(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtPreise_Change")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Select Case Index
    Case 2, 10, 11, 13
        Call CalcGesamtBrutto
End Select

Call DefErrPop
End Sub

Private Sub txtPreise_GotFocus(Index As Integer)
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

With txtPreise(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtXML_DblClick(Index As Integer)

If (Index >= 3) Then
    With wbXML
        .Left = txtXML(0).Left
        .Top = txtXML(0).Top
        .Width = picStammdatenBack(6).Width - .Left - 150
        .Height = txtXML(5).Top + txtXML(5).Height - .Top
        
        Dim iFile%
        Dim sDatei$
        sDatei = CurDir() + "\eRezeptXML.xml"
        iFile = FreeFile
        Open sDatei For Output As #iFile
        Print #iFile, txtXML(Index).text
        Close #iFile
        .Navigate (sDatei)
        
        .Visible = True
    End With
End If

End Sub

Sub CheckLieferengpass()
Dim i%

'MsgBox ("check")
lblLieferengpass.Visible = False
For i% = 2 To 3
    With cboZA_Schluessel(i - 1)
        If (.ListIndex = 2) Then
'MsgBox ("check: " + CStr(i - 1) + " " + CStr(.ListIndex))
            lblLieferengpass.Visible = True
            Exit For
        End If
    End With
Next i%

End Sub

Sub CheckFreitext(Index As Integer)
Dim i%

If (Index >= 1) And (Index <= 3) Then
    With cboZA_Schluessel(Index)
        If (.ListIndex = 2) Or (.ListIndex = 4) Then
            If (txtZA_Freitext(Index).text = "") Then
                Dim sFreitext As String
                For i = 1 To 3
                    sFreitext = Trim(txtZA_Freitext(i).text)
                    If (sFreitext <> "") Then
                        Exit For
                    End If
                Next i
                txtZA_Freitext(Index).text = sFreitext
            End If
        End If
    End With
End If

End Sub

Private Sub cmdDruckVerordnung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdDruckVerordnung_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim h$, h2$, header2$, sDruckerOrg$

Dim hPrinter As Printer
sDruckerOrg = Printer.DeviceName
For Each hPrinter In Printers
    If (UCase(StandardDrucker) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next

'Printer.Orientation = vbPRORLandscape   ' vbPRORPortrait

'Call StartAnimation(Me, "Ausdruck wird erstellt ...")

wbVerordnung.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
'you can change OLECMDEXECOPT_DONTPROMPTUSER to OLECMDEXECOPT_PROMPTUSER if you wish

For Each hPrinter In Printers
    If (UCase(sDruckerOrg) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next
              
'Call StopAnimation(Me)
Call MessageBox("Ausdruck wurde erstellt !", vbInformation)

Call DefErrPop
End Sub

Private Function FormatOrEmpty(dVal#) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FormatOrEmpty")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

FormatOrEmpty = IIf(dVal > 0, Format(dVal, "0.00"), "")

Call DefErrPop
End Function

Private Sub CalcGesamtBrutto()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcGesamtBrutto")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With txtPreise(1)
    .text = FormatOrEmpty(xVal(txtPreise(2).text) + xVal(txtPreise(10).text) + xVal(txtPreise(11).text) + xVal(txtPreise(13).text))
End With

Call DefErrPop
End Sub

Private Sub txtZA_Freitext_Change(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtZA_Freitext_Change")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Index = 15) Then
    Dim pzn$, txt$
    
    pzn = txtZA_Freitext(15).text
    chkZA_Gruppe(15).Value = Abs(pzn <> "")
    lblSpenderPzn.Caption = ""
    
    If (pzn <> "") Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If Not (TaxeRec.EOF) Then
            txtZA_Freitext(15).text = PznString(TaxeRec!pzn)
        
            txt = Left$(TaxeRec!Name + Space$(29), 29) + Mid$(TaxeRec!menge, 3) + TaxeRec!einheit
            lblSpenderPzn.Caption = txt
        End If
        chkZA_Gruppe(15).Value = Abs(Not (TaxeRec.EOF))
        TaxeRec.Close
    End If
End If

Call DefErrPop
End Sub

'Statistiksatz für Artikel:PZN$ aufbereiten MENGE% = Liefermenge -------
Sub WuBuchen(pzn$, menge%, Optional BenutzerNr% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuBuchen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim i%, asatz%, ssatz%, LagAlt%, LagNeu%
Dim lager#

asatz% = 0
If (Val(pzn$) <> 0) And (Val(pzn$) <> 9999999) Then
    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
    If (FabsErrf% = 0) Then
        asatz% = 1  'Abs(ArtikelAdoRec!LagerKz)
        ssatz = 1
    End If
End If

If (FabsErrf% <> 0) Or (asatz% = 0) Or (Val(pzn$) = 9999999) Or (menge% = 0) Then
    Call DefErrPop
    Exit Sub
End If

If (FabsErrf% = 0) Then
    LagAlt% = ass.poslag
    
    lager# = CSng(ass.poslag) + CSng(menge%)
    If (lager# > 32000) Then lager# = 32000
    If (lager# < -32000) Then lager# = -32000
    ass.poslag = CInt(lager#)
    
    SQLStr = "UPDATE Artikel SET"
    SQLStr = SQLStr + " Poslag=" + CStr(ass.poslag)
    SQLStr = SQLStr + ", LagerBew='" + SqlString(ass.lagbew) + "'"
    SQLStr = SQLStr + " WHERE Pzn=" + pzn$
    Call SqlExecute(Artikel, SQLStr)
    
    LagNeu% = ass.poslag
    Call LagerBew(pzn$, LagAlt%, LagNeu%, BenutzerNr%, 1)
End If

Call DefErrPop
End Sub

Function SqlExecute%(SqlDB As Object, SQLStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlExecute%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErr2:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul, 9999)
Case vbRetry
  Resume
End Select
Call DefErrPop: Exit Function
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%, SqlLog%
Dim ErrorNumber&, lRecs&
Dim s$

SqlLog = FreeFile
Open "SQLOP.LOG" For Append As #SqlLog%
s$ = Format(Now, "dd.mm.yyyy") + " " + Format(Now, "hh:nn:ss") + "     " + SQLStr
Print #SqlLog%, s

lRecs = 0
On Error Resume Next
Call SqlDB.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
ErrorNumber = Err.Number
Err.Clear
s = "Erg: " + CStr(lRecs)
Print #SqlLog%, s
If (ErrorNumber <> 0) Then
    s = "Err: " + CStr(ErrorNumber)
    Print #SqlLog%, s
    Call DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul, 9999)
End If

Close #SqlLog%

SqlExecute = lRecs

Call DefErrPop
End Function


Sub LagerBew(pzn$, LagAlt%, LagNeu%, BenutzerNr%, Optional ProgNr% = 7)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LagerBew")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Zeiger!, Groesse!
Dim s$
Dim Lb1 As clsLagerBewegung
Dim lbdb1 As clsLagerbewegungDB

If (LagAlt% <> LagNeu%) And (InStr(para.Benutz, "V") > 0) Then
    Set Lb1 = New clsLagerBewegung
    Set lbdb1 = New clsLagerbewegungDB
    If (lbdb1.DBvorhanden) Then
        lbdb1.OpenDB
        
        s$ = Format(Now, "DD.MM.YYYY HH:MM")
        SQLStr = "INSERT INTO Lagerbewegungen (Pzn,Datum,LagAlt,LagNeu,Computer,Programm,Benutzer)"
        SQLStr = SQLStr + " VALUES (" + pzn$ + ",'" + s + "'," + CStr(LagAlt) + "," + CStr(LagNeu) + "," + CStr(Val(para.User)) + "," + CStr(ProgNr) + "," + CStr(BenutzerNr%)
        SQLStr = SQLStr + ")"
        Call SqlExecute(lbdb1, SQLStr)
'        Call lbdb1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
        
    '  f.tmp% = 6
    '  If ProgrammNr$ = "" Then
    '    Open "MPOS.DAT" For Random Access Read Write Shared As #f.tmp% Len = 100
    '    FIELD #f.tmp%, 2 AS F.PWCODE$, 2 AS F.Z$, 2 AS F.M$, 20 AS F.N$,2 AS F.FCOMP$, 2 AS F.PRG$
    '    GET #f.tmp%, VAL(user$) + 1
    '    werwardas$ = Chr$(CVI(f.PWCODE$))
    '    Close #f.tmp%
    '    ProgrammNr$ = Chr$(0)
    '    '1 VERKAUF
    '    '2 INFOART
    '    '3 STAMA
    '    '4 VERKERF
    '    '5 POSDRUCK
    '    '6 FAKT
    '    '7 WARENEIN
    '    ProgrammNr$ = Chr$(7)
    '  End If
    
        lbdb1.CloseDB
    End If
End If

Call DefErrPop
End Sub



