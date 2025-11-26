VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmMatchcodeKunden 
   Caption         =   "Matchcode - Auswahl"
   ClientHeight    =   5715
   ClientLeft      =   255
   ClientTop       =   600
   ClientWidth     =   11655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   11655
   Begin VB.ListBox lstSortierung 
      Height          =   255
      Left            =   10320
      Sorted          =   -1  'True
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDatei 
      Height          =   375
      Index           =   0
      Left            =   10560
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox picSave 
      Height          =   1095
      Left            =   10440
      ScaleHeight     =   1035
      ScaleWidth      =   915
      TabIndex        =   12
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
      TabIndex        =   10
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
         TabIndex        =   11
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   450
         Index           =   0
         Left            =   3360
         TabIndex        =   4
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1200
      End
      Begin VB.TextBox txtMatchcode 
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid flxInfo 
         Height          =   1500
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   840
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   2646
         _Version        =   65541
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
         Height          =   420
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         _Version        =   65541
         Rows            =   0
         Cols            =   0
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
      Begin VB.TextBox txtFlexBack 
         Height          =   615
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2520
         Width           =   1935
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
         TabIndex        =   9
         Top             =   240
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   270
         Width           =   9615
      End
   End
   Begin ComctlLib.ImageList imgToolbar 
      Index           =   0
      Left            =   10320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":03A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":04B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":05C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":085A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":0AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":0BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":0D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1262
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1374
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1486
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1598
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":16AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":17BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":18CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":19E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1D16
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
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":20BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":25DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":2870
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":2B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":2D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":3026
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":32B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":3B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":3D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":402E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":42C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":4552
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":47E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":4A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":4D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":522C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMatchcodeKunden.frx":54BE
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
   End
End
Attribute VB_Name = "frmMatchcodeKunden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const INI_DATEI = "\user\winop.ini"
'Const INI_SECTION = "Matchcode"
'Const INFO_SECTION = "Matchcode Infobereich"

Dim WithEvents opToolbar As clsToolbar
Attribute opToolbar.VB_VarHelpID = -1
Dim opBereich As clsOpBereiche
Dim InfoMain As clsInfoBereich

Dim HochfahrenAktiv%

Dim Standard%

Dim FabsErrf%
Dim FabsRecno&


Private Const DefErrModul = "frmMatchcodeKunden.frm"

Private Sub cmdDatei_Click(index As Integer)
Dim erg%

MatchModus% = index
erg% = Match1.SuchArtikel%(OrgSuch$, opBereich.ArbeitAnzZeilen)

End Sub

Private Sub cmdEsc_Click(index As Integer)
Unload Me
End Sub

Private Sub cmdOk_Click(index As Integer)
Dim ind%, erg%, row%, col%
Dim h$, h2$, pzn$, txt$

If (ActiveControl.Name = txtMatchcode.Name) Then
    h$ = RTrim(UCase(txtMatchcode.text))
    erg% = Match1.SuchArtikel%(h$, opBereich.ArbeitAnzZeilen)
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
        erg% = Match1.SuchArtikel%(h$, opBereich.ArbeitAnzZeilen)
    Else
        row% = flxarbeit(0).row
        MatchcodePzn$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
        With flxarbeit(0)
            MatchcodeTxt$ = Trim$(.TextMatrix(.row, 1)) + "  " + Trim$(.TextMatrix(.row, 2)) + " " + Trim$(.TextMatrix(.row, 3))
        End With
        Unload Me
    End If
ElseIf (ActiveControl.Name = flxInfo(0).Name) Then
    With flxInfo(0)
        row% = .row
        col% = .col
        h$ = RTrim(.text)
    End With
    If (col% = 0) Then
        If (MatchTyp% = MATCH_ARTIKEL) Then
            If (h$ = TEXT_SONDERANGEBOTE) Then
                With flxarbeit(0)
                    pzn$ = Format$(Ausgabe(flxarbeit(0).row - 1).pzn, "0000000")
                    h2$ = RTrim$(.TextMatrix(.row, 1)) + "  " + RTrim$(.TextMatrix(.row, 2))
                    h2$ = h2$ + RTrim$(.TextMatrix(.row, 3))
                End With
                Call clsDialog.ZeigeAngebote(pzn$, h2$, 0, 0)
            ElseIf (h$ = TEXT_BESTELLSTATUS) Then
                With flxarbeit(0)
                    pzn$ = Format$(Ausgabe(flxarbeit(0).row - 1).pzn, "0000000")
                    h2$ = RTrim$(.TextMatrix(.row, 1)) + "  " + RTrim$(.TextMatrix(.row, 2))
                    h2$ = h2$ + RTrim$(.TextMatrix(.row, 3))
                End With
                Call clsDialog.BestellStatus(pzn$, h2$)
            ElseIf (h$ = TEXT_STATISTIK) Then
                pzn$ = Format$(Ausgabe(flxarbeit(0).row - 1).pzn, "0000000")
                Call clsDialog.ZeigeStatbild(pzn$, Me)
                AppActivate Me.Caption
            End If
        ElseIf (MatchTyp% = MATCH_LIEFERANTEN) Then
            If (h$ = TEXT_ABSAGEN) Or (h$ = TEXT_NACHBEARBEITUNG) Then
                row% = flxarbeit(0).row
                ind% = Ausgabe(row% - 1).pzn
                txt$ = flxarbeit(0).TextMatrix(row%, 1)
                Call clsDialog.AnzeigeFenster(h$, ind%, txt$)
            End If
        End If
    ElseIf (col% Mod 2) Then
        Call InfoMain.EditInfoBelegung
        Call AuswahlKurzInfo
    End If
End If

End Sub

Private Sub cmdToolbar_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index <= 8) Then
    Call mnuBearbeitenInd_Click(index - 1)
ElseIf (index <= 16) Then
    Call mnuBearbeitenInd_Click(index)
ElseIf (index = 19) Then
'    Call mnuBeenden_Click
End If

End Sub

Private Sub flxarbeit_DblClick(index As Integer)
cmdOk(0).Value = True
End Sub

Private Sub flxarbeit_GotFocus(index As Integer)
txtFlexBack.SetFocus
End Sub

Private Sub flxarbeit_RowColChange(index As Integer)
    
If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

If ((flxarbeit(0).Visible = True) And (KeinRowColChange% = False)) Then
    Call AuswahlKurzInfo
    flxInfo(0).row = 0
    flxInfo(0).col = 0
End If

End Sub

Private Sub flxInfo_DblClick(index As Integer)
cmdOk(0).Value = True
End Sub

Private Sub flxInfo_GotFocus(index As Integer)
Dim i%, ArbeitRow%, InfoRow%, aRow%, aCol%
Dim pzn$

If (index = 0) Then
    ArbeitRow% = flxarbeit(0).row
    pzn$ = Format$(Ausgabe(ArbeitRow% - 1).pzn, "0000000")
    
    With flxInfo(0)
        aRow% = .row
        aCol% = .col
        
        InfoRow% = 0
        
        If (MatchTyp% = MATCH_ARTIKEL) Then
            If (clsDialog.TesteBstatus(pzn$)) Then
                .TextMatrix(InfoRow%, 0) = TEXT_BESTELLSTATUS
                InfoRow% = InfoRow% + 1
            End If
            
            GhAngebotMax& = clsAngebote1.DateiLen / 34 - 1
            If (AngebotSuchen&(pzn$) > 0) Then
                .TextMatrix(InfoRow%, 0) = TEXT_SONDERANGEBOTE
                InfoRow% = InfoRow% + 1
            End If
            
            If (MatchModus% = LAGER_MATCH) Or (Ausgabe(ArbeitRow% - 1).LagerKz = 2) Then
                .TextMatrix(InfoRow%, 0) = TEXT_STATISTIK
                InfoRow% = InfoRow% + 1
            End If
        ElseIf (MatchTyp% = MATCH_LIEFERANTEN) Then
                .TextMatrix(InfoRow%, 0) = TEXT_ABSAGEN
                InfoRow% = InfoRow% + 1
                .TextMatrix(InfoRow%, 0) = TEXT_NACHBEARBEITUNG
                InfoRow% = InfoRow% + 1
        End If
        
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

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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

End Sub

Private Sub Form_Load()
Dim i%
Dim l&
Dim h$

HochfahrenAktiv% = True
   
'Top = frmAction.Top + 600
'Left = frmAction.Left + 600
'Width = frmAction.Width - 1200
'Height = frmAction.Height - 1200
Top = 0
Left = 0
Width = Screen.Width - 1200
Height = Screen.Height - 1200

Caption = Match1.IniSection

With picSave
    .Left = 0
    .Top = 0
    .Width = ScaleWidth
    .Height = ScaleHeight
    .ZOrder 0
End With


h$ = "0"
l& = GetPrivateProfileString(Match1.IniSection, "Standard", "0", h$, 2, INI_DATEI)
Standard% = Val(Left$(h$, l&))
MatchModus% = Standard%



Set opToolbar = New clsToolbar
Call opToolbar.InitToolbar(Me, INI_DATEI, Match1.IniSection$)

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
cmdToolbar(14).ToolTipText = "shift+F7"
cmdToolbar(15).ToolTipText = "shift+F8"
cmdToolbar(16).ToolTipText = "shift+F9"
'cmdToolbar(19).ToolTipText = "Programm beenden"



Call wPara1.InitFont(Me)

Set InfoMain = New clsInfoBereich
Call InfoMain.InitInfoBereich(flxInfo(0), INI_DATEI, Match1.InfoSection$)
Call InfoMain.ZeigeInfoBereich("", False)
flxInfo(0).row = 0
flxInfo(0).col = 0

Set opBereich = New clsOpBereiche
Call opBereich.InitBereich(Me, opToolbar)
opBereich.ArbeitTitel = False
opBereich.ArbeitLeerzeileOben = True
opBereich.ArbeitWasDarunter = False
opBereich.InfoTitel = False
opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
opBereich.AnzahlButtons = -2

Call InitDateiButtons

If (MatchcodeTxt$ <> "") Then
    txtMatchcode.text = MatchcodeTxt$
End If

HochfahrenAktiv% = False

End Sub

Sub RefreshBereichsFlexSpalten()
Dim i%, j%, spBreite%
Dim sp&
            
Call Match1.MachAuswahlGrid(opBereich.ArbeitAnzZeilen)
With flxInfo(0)
    sp& = .Width / 8
    .ColWidth(0) = 2 * sp&
    For i% = 1 To 6
        .ColWidth(i%) = sp&
    Next i%
End With
   
End Sub

Sub RefreshBereichsControlsAdd()
Dim i%

On Error Resume Next

ReDim Ausgabe(opBereich.ArbeitAnzZeilen - 1)
lblMatchcode.Left = wPara1.LinksX
lblMatchcode.Top = wPara1.TitelY   'FlexY%
txtMatchcode.Left = lblMatchcode.Left + lblMatchcode.Width + 150
txtMatchcode.Top = lblMatchcode.Top
txtFlexBack.Top = flxarbeit(0).Top + 15
txtFlexBack.Left = flxarbeit(0).Left
End Sub

Sub RefreshBereichsFarbenAdd()
Dim i%

On Error Resume Next

lblMatchcode.BackColor = wPara1.FarbeArbeit

End Sub

Sub MachMaske()
Dim i%, j%, erg%, h$, ind%, ind2%, sAnz%, NurLagerndeAktiv%

AnzAnzeige% = opBereich.ArbeitAnzZeilen - 1
For i% = 1 To (AnzAnzeige% - 1)
    erg% = Match1.SuchWeiter%(i% - 1, True)
    If (erg%) Then
        Call Match1.Umspeichern(buf$, i%)
    Else
        AnzAnzeige% = i%
        Exit For
    End If
Next i%

Call AuswahlBefüllen
flxarbeit(0).row = 1

End Sub

Sub AuswahlBefüllen()
Dim i%, j%, k%, ind%, AltRow%
Dim Suc&
Dim h$

On Error Resume Next

With flxarbeit(0)
    KeinRowColChange% = True

    txtFlexBack.Visible = False
    .Visible = False
    AltRow% = .row
    
    For i% = 1 To AnzAnzeige%
        
        h$ = Ausgabe(i% - 1).Name
        For j% = 0 To .Cols - 2
            ind% = InStr(h$, vbTab)
            .TextMatrix(i%, j%) = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Next j%
        .TextMatrix(i%, .Cols - 1) = h$
    Next i%
        
    .row = 1
    .col = .Cols - 1
    h$ = .CellFontName
    
    
    .FillStyle = flexFillRepeat
    For i% = 1 To AnzAnzeige%
        .row = i%
        .col = 0
        .ColSel = .Cols - 1
        
        If (Ausgabe(i% - 1).LagerKz = 2) Then
            .CellFontBold = True
        Else
            .CellFontBold = False
        End If
    Next i%
    .FillStyle = flexFillSingle
    
    .Rows = AnzAnzeige% + 1
    
    If (MatchModus% = LAGER_MATCH) Then
        For i% = 1 To AnzAnzeige%
            .row = i%
            .col = .Cols - 1
            .CellFontName = "Courier New"
        Next i%
    End If
    
    .row = AltRow%
    .col = 0
    .Visible = True
    txtFlexBack.Visible = True
    Call AuswahlKurzInfo
    KeinRowColChange% = False
End With
End Sub

Public Sub AuswahlKurzInfo()
Dim row%, iRow%, iCol%
Dim pzn$, ch$, ActKontrollen$, ActZusatz$, actzuordnung$

row% = flxarbeit(0).row

ActKontrollen$ = ""
actzuordnung$ = ""
ActZusatz$ = ""

pzn$ = Ausgabe(row% - 1).pzn

iRow% = flxInfo(0).row
iCol% = flxInfo(0).col
Call InfoMain.ZeigeInfoBereich(pzn$)
Call ZeigeInfoBereichAdd(0)
flxInfo(0).row = iRow%
flxInfo(0).col = iCol%

End Sub

Sub ZeigeInfoBereichAdd(index%)
Dim j%

With flxInfo(index%)
    .Redraw = False
    
    .row = 0
    .col = 0
    .CellFontBold = True
    .TextMatrix(0, 0) = MatchAnzeigeTyp$(MatchModus%)
    
    j% = 2
    Do While (j% <= InfoMain.AnzInfoZeilen%)
        .TextMatrix(j% - 1, 0) = ""
        j% = j% + 1
    Loop
    
    .Redraw = True
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call opToolbar.SpeicherIniToolbar
Set opToolbar = Nothing
Set InfoMain = Nothing
End Sub

Private Sub mnuBearbeitenInd_Click(index As Integer)
Dim erg%, row%, col%, ind%
Dim l&
Dim pzn$, txt$

Select Case index

    Case MENU_F2
        If (ActiveControl.Name = flxInfo(0).Name) Then
            Call InfoMain.InsertInfoBelegung(flxInfo(0).row)
            opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
            Call opBereich.RefreshBereich
            Call AuswahlKurzInfo
        ElseIf (MatchTyp% = MATCH_ARTIKEL) Then
            MatchcodePzn$ = "9999999"
            MatchcodeTxt$ = Trim$(txtMatchcode.text)
            Unload Me
        End If
    
    Case MENU_F3
        erg% = clsDialog.WechselFenster(MatchAnzeigeTyp$, Standard%)
        l& = WritePrivateProfileString(Match1.IniSection$, "Standard", Str$(Standard%), INI_DATEI)
        If (erg% >= 0) Then
            cmdDatei(erg%).Value = True
        End If
        
    Case MENU_F5
        If (ActiveControl.Name = flxInfo(0).Name) Then
            Call InfoMain.LoescheInfoBelegung(flxInfo(0).row, (flxInfo(0).col - 1) \ 2)
            opBereich.InfoAnzZeilen = InfoMain.AnzInfoZeilen
            Call opBereich.RefreshBereich
            Call AuswahlKurzInfo
        End If
    
    Case MENU_F8
        If (ActiveControl.Name = flxInfo(0).Name) Then
            col% = flxInfo(0).col
            If (col% > 0) And (col% Mod 2) Then
                row% = flxInfo(0).row
                If (InfoMain.Bezeichnung(row%, (col% - 1) \ 2) <> "") Then
                    Call EditInfoName
                End If
            End If
        Else
            With flxarbeit(0)
                If (RTrim$(.TextMatrix(.row, 1)) <> "") Then
                    If (MatchTyp% = MATCH_ARTIKEL) Then
                        pzn$ = Ausgabe(.row - 1).pzn
                        txt$ = RTrim$(.TextMatrix(.row, 1)) + " " + RTrim$(.TextMatrix(.row, 2)) + RTrim$(.TextMatrix(.row, 3))
                        Call clsDialog.ZusatzFenster(ZUSATZ_ARTIKEL, pzn$, txt$)
                    Else
                        ind% = Ausgabe(.row - 1).pzn
                        txt$ = RTrim$(.TextMatrix(.row, 1))
                        Call clsDialog.ZusatzFenster(ZUSATZ_LIEFERANTEN, ind%, txt$)
                    End If
                End If
            End With
        End If
End Select

End Sub

Private Sub mnuBeenden_Click()
Unload Me
End Sub

Private Sub txtFlexBack_GotFocus()
flxarbeit(0).HighLight = flexHighlightAlways
End Sub

Private Sub txtFlexBack_lostFocus()
flxarbeit(0).HighLight = flexHighlightNever
End Sub

Private Sub txtFlexBack_KeyDown(KeyCode As Integer, Shift As Integer)
Dim erg%

Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
        Call AuswahlRowChange(KeyCode)
        KeyCode = 0
End Select

flxarbeit(0).ColSel = flxarbeit(0).Cols - 1

End Sub

Sub AuswahlRowChange(KeyCode As Integer)
Dim erg%, i%, j%, h$, ind%, neu%, NurLagerndeAktiv%, sAnz%
    
neu% = True
With flxarbeit(0)
    Select Case KeyCode
        Case vbKeyUp
            If (.row > 1) Then
                .row = .row - 1
                neu% = False
            Else
                erg% = Match1.SuchWeiter%(0, False)
                If (erg%) Then
                    GoSub ZeileHinein
                Else
                    .row = 1
                End If
            End If
        Case vbKeyPageUp
            erg% = Match1.SuchWeiter%(0, False)
            If (erg%) Then
                For j% = 1 To (opBereich.ArbeitAnzZeilen - 1)
                    GoSub ZeileHinein
                    erg% = Match1.SuchWeiter%(0, False)
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
                erg% = Match1.SuchWeiter%(AnzAnzeige% - 1, True)
                If (erg%) Then
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    Call Match1.Umspeichern(buf$, AnzAnzeige% - 1)
                Else
                    .row = AnzAnzeige%
                End If
            End If
        Case vbKeyPageDown
            erg% = Match1.SuchWeiter%(AnzAnzeige% - 1, True)
            If (erg%) Then
                For j% = 1 To AnzAnzeige%
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    Call Match1.Umspeichern(buf$, AnzAnzeige% - 1)
                    erg% = Match1.SuchWeiter%(AnzAnzeige% - 1, True)
                    If (erg% = False) Then Exit For
                Next j%
            Else
                .row = AnzAnzeige%
            End If
    End Select
End With


If (neu% = True) Then Call AuswahlBefüllen
txtFlexBack.SetFocus
Exit Sub

ZeileHinein:
For i% = (opBereich.ArbeitAnzZeilen - 2) To 1 Step -1
    Ausgabe(i%) = Ausgabe(i% - 1)
Next i%
Call Match1.Umspeichern(buf$, 0)
If (AnzAnzeige% < (opBereich.ArbeitAnzZeilen - 1)) Then
    AnzAnzeige% = AnzAnzeige% + 1
    flxarbeit(0).Rows = AnzAnzeige% + 1
End If
Return

End Sub

Private Sub txtMatchcode_GotFocus()

With txtMatchcode
    .SelStart = 0
    .SelLength = Len(.text)
End With

End Sub

Private Sub mnuToolbarGross_Click()
Dim i%

If (opToolbar.BigSymbols) Then
    opToolbar.BigSymbols = False
Else
    opToolbar.BigSymbols = True
End If

End Sub

Private Sub mnuToolbarLabels_Click()

If (opToolbar.Labels) Then
    opToolbar.Labels = False
Else
    opToolbar.Labels = True
End If

End Sub

Private Sub mnuToolbarPositionInd_Click(index As Integer)
Dim i%

opToolbar.Position = index

End Sub

Private Sub mnuToolbarVisible_Click()

If (opToolbar.Visible) Then
    opToolbar.Visible = False
    mnuToolbarVisible.Caption = "Einblenden"
Else
    opToolbar.Visible = True
    mnuToolbarVisible.Caption = "Ausblenden"
End If

End Sub

Private Sub opToolbar_Resized()
Call opBereich.ResizeWindow
End Sub

Private Sub Form_Resize()

On Error Resume Next

If (HochfahrenAktiv%) Then Exit Sub

If (Me.WindowState = vbMinimized) Then Exit Sub

Call opBereich.ResizeWindow
picSave.Visible = False

End Sub

Private Sub flxarbeit_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call opToolbar.Move(flxarbeit(index), picBack(index), Source, X, Y)
End Sub

Private Sub flxInfo_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call opToolbar.Move(flxInfo(index), picBack(index), Source, X, Y)
End Sub

Private Sub lblarbeit_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call opToolbar.Move(lblArbeit(index), picBack(index), Source, X, Y)
End Sub

Private Sub lblInfo_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call opToolbar.Move(lblInfo(index), picBack(index), Source, X, Y)
End Sub

Private Sub picBack_DragDrop(index As Integer, Source As Control, X As Single, Y As Single)
Call opToolbar.Move(picBack(index), picBack(index), Source, X, Y)
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picToolbar.Drag (vbBeginDrag)
opToolbar.DragX = X
opToolbar.DragY = Y
End Sub

Sub EditInfoName()
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
    InfoMain.Bezeichnung(EditRow%, (EditCol% - 1) \ 2) = EditTxt$
    Call AuswahlKurzInfo
End If

End Sub

Sub InitDateiButtons()
Dim i%, Max%

Max% = UBound(MatchAnzeigeTyp)
For i% = 1 To Max%
    Load mnuDateiInd(i%)
    Load cmdDatei(i%)
Next i%

For i% = 0 To Max%
    cmdDatei(i%).Top = 0
    cmdDatei(i%).Left = i% * 900
    cmdDatei(i%).Visible = True
    cmdDatei(i%).ZOrder 1
Next i%

If (MatchTyp% = MATCH_ARTIKEL) Then
    mnuDateiInd(0).Caption = "&Taxe"
    mnuDateiInd(1).Caption = "Taxe &phonetisch"
    mnuDateiInd(2).Caption = "&Lagerartikel"

    cmdDatei(0).Caption = "&T"
    cmdDatei(1).Caption = "&P"
    cmdDatei(2).Caption = "&L"
ElseIf (MatchTyp% = MATCH_LIEFERANTEN) Then
    mnuDateiInd(0).Caption = "&Lieferanten"
    cmdDatei(0).Caption = "&L"
End If

End Sub
