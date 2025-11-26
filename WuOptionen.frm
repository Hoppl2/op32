VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWuOptionen 
   Caption         =   "WÜ-Optionen"
   ClientHeight    =   7620
   ClientLeft      =   -15
   ClientTop       =   585
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8610
   Begin VB.CommandButton cmdAMPV 
      Caption         =   "&AMPV"
      Height          =   450
      Left            =   7440
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Hinunter (>)"
      Height          =   450
      Index           =   1
      Left            =   8400
      TabIndex        =   15
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Hinauf (<)"
      Height          =   450
      Index           =   0
      Left            =   8520
      TabIndex        =   14
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtOptionen 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstOptionen 
      Height          =   450
      Left            =   6000
      TabIndex        =   4
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
      TabIndex        =   17
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   16
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   6765
      Left            =   -1560
      TabIndex        =   18
      Top             =   120
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   11933
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
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
      TabCaption(0)   =   "&1 - Tätigkeiten"
      TabPicture(0)   =   "WuOptionen.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeWuOptionen(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Aufschlagstab."
      TabPicture(1)   =   "WuOptionen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeWuOptionen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Preiserstellung"
      TabPicture(2)   =   "WuOptionen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeWuOptionen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Sortierung"
      TabPicture(3)   =   "WuOptionen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeWuOptionen(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5 - Verfall"
      TabPicture(4)   =   "WuOptionen.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeWuOptionen(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Sonstiges"
      TabPicture(5)   =   "WuOptionen.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "fmeWuOptionen(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   4
         Left            =   -73320
         TabIndex        =   41
         Top             =   2280
         Width           =   7455
         Begin MSFlexGridLib.MSFlexGrid flxOptionen 
            Height          =   1560
            Index           =   4
            Left            =   2280
            TabIndex        =   42
            Top             =   480
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   2752
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
      End
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   5175
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   720
         Width           =   8415
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Abfrage &Lieferant bei man. Erfassung"
            Height          =   375
            Index           =   10
            Left            =   5880
            TabIndex        =   30
            Top             =   3360
            Width           =   3615
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Nicht-&Rezeptpfl. AM zur Kalkulation anbieten"
            Height          =   375
            Index           =   9
            Left            =   5520
            TabIndex        =   29
            Top             =   3000
            Width           =   3615
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Bestände berücksichtigen bei (Partner-)Laden..."
            Height          =   375
            Index           =   8
            Left            =   4320
            TabIndex        =   27
            Top             =   2520
            Width           =   3615
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Teilbestellungen bei (Partner-)Ladenhüterbest."
            Height          =   375
            Index           =   7
            Left            =   4680
            TabIndex        =   25
            Top             =   2040
            Width           =   3615
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Senden von jedem Plat&z"
            Height          =   375
            Index           =   6
            Left            =   4440
            TabIndex        =   23
            Top             =   1560
            Width           =   3615
         End
         Begin VB.ComboBox cboSonstiges 
            Height          =   315
            Index           =   1
            Left            =   4200
            Style           =   2  'Dropdown-Liste
            TabIndex        =   38
            Top             =   4800
            Width           =   3255
         End
         Begin VB.ComboBox cboSonstiges 
            Height          =   315
            Index           =   0
            Left            =   4320
            Style           =   2  'Dropdown-Liste
            TabIndex        =   36
            Top             =   4320
            Width           =   3255
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Pharmabox in D&OS"
            Height          =   375
            Index           =   5
            Left            =   4560
            TabIndex        =   21
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtSonstiges 
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
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   34
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3720
            Width           =   495
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Besorgerinfo muß durch F4 quittiert werden"
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   28
            Top             =   2520
            Width           =   5775
         End
         Begin VB.TextBox txtSonstiges 
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
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   32
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   3120
            Width           =   495
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Teil&defekte bei Lagerartikel automatisch wiederbestellen"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   5775
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "Bestellvorschlags-&Protokoll erstellen"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   5775
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Lagerkontroll-Liste drucken"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Width           =   5775
         End
         Begin VB.CheckBox chkSonstiges 
            Caption         =   "&Etiketten bereitstellen"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   20
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lblSonstiges 
            Caption         =   "Nach &man. LM-Bestätigung springen zu"
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
            Index           =   3
            Left            =   -120
            TabIndex        =   37
            Top             =   4680
            Width           =   3975
         End
         Begin VB.Label lblSonstigesRechts 
            Caption         =   "Tage speichern"
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
            Left            =   5640
            TabIndex        =   40
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label lblSonstiges 
            Caption         =   "&Automatische Ausdrucke an"
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
            Left            =   0
            TabIndex        =   35
            Top             =   4200
            Width           =   3975
         End
         Begin VB.Label lblSonstigesRechts 
            Caption         =   "x"
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
            Left            =   5760
            TabIndex        =   39
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label lblSonstiges 
            Caption         =   "&Anzahl  Retouren-Ausdruck"
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
            Left            =   240
            TabIndex        =   33
            Top             =   3720
            Width           =   3975
         End
         Begin VB.Label lblSonstiges 
            Caption         =   "&Sendeprotokolle der letzten"
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
            Left            =   0
            TabIndex        =   31
            Top             =   3120
            Width           =   3975
         End
      End
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   2775
         Index           =   3
         Left            =   -74280
         TabIndex        =   11
         Top             =   720
         Width           =   6855
         Begin MSFlexGridLib.MSFlexGrid flxOptionen 
            Height          =   1080
            Index           =   3
            Left            =   1680
            TabIndex        =   2
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1905
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
      End
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   2775
         Index           =   2
         Left            =   -74520
         TabIndex        =   9
         Top             =   720
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid flxOptionen 
            Height          =   1320
            Index           =   2
            Left            =   2160
            TabIndex        =   10
            Top             =   960
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2328
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
      End
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   2415
         Index           =   1
         Left            =   -74400
         TabIndex        =   8
         Top             =   840
         Width           =   6975
         Begin MSFlexGridLib.MSFlexGrid flxOptionen 
            Height          =   1560
            Index           =   1
            Left            =   2280
            TabIndex        =   1
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2752
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
      End
      Begin VB.Frame fmeWuOptionen 
         Caption         =   "Frame1"
         Height          =   3015
         Index           =   0
         Left            =   -74400
         TabIndex        =   7
         Top             =   840
         Width           =   7215
         Begin MSFlexGridLib.MSFlexGrid flxOptionen 
            Height          =   720
            Index           =   0
            Left            =   1800
            TabIndex        =   0
            Top             =   1200
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1270
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
      End
   End
End
Attribute VB_Name = "frmWuOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "WUOPTIONEN.FRM"

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
Dim i%, start%, hTab%

hTab% = tabOptionen.Tab
With flxOptionen(hTab%)
    start% = 0
    If (hTab% = 1) Then start% = 1
    For i% = start% To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

Call DefErrPop
End Sub

Private Sub cmdAMPV_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAMPV_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, start%, hTab%

hTab% = tabOptionen.Tab
With flxOptionen(hTab%)
    .TextMatrix(.row, 2) = "AMPV"
    .SetFocus
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
    Call SpeicherIniTaetigkeiten
    Call SpeicherIniAufschlagsTabelle
    Call SpeicherIniRundungen
    Call SpeicherIniWuSortierungen
    Call SpeicherIniVerfallWarnungen
    Call frmAction.SpeicherIniWerte

    Call PruefeTaetigkeiten
    
    OptionenModus% = 2
    OptionenNeu% = True
    Unload Me

Else
    hTab% = tabOptionen.Tab
    If (hTab% < 5) Then
        row% = flxOptionen(hTab%).row
        col% = flxOptionen(hTab%).col
        If (lstOptionen.Visible = True) Then
            h$ = lstOptionen.text
            lstOptionen.Visible = False
        ElseIf (txtOptionen.Visible = True) Then
            h$ = RTrim(txtOptionen.text)
            txtOptionen.Visible = False
            tabOptionen.Enabled = True
            With flxOptionen(hTab%)
                .TextMatrix(.row, .col) = h$
                .SetFocus
                If (.col < .Cols - 2) Then .col = .col + 1
                If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
            End With
        Else
            If (hTab% = 1) And (row% = MAX_AUFSCHLAEGE) Then
                Call iMsgBox("Falls Sie in dieser Zeile Werte eintragen, werden diese zur Kalkulation der Besorger herangezogen !", vbInformation)
            End If
            
            s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
            
            Select Case flxOptionen(hTab%).col
                Case 0
                    If (hTab% = 0) Or (hTab% = 2) Or (hTab% = 3) Then
                        Call EditOptionenLst
                    ElseIf (hTab% = 4) Then
                        Call EditOptionenTxt
                    End If
                Case 1
                    If (hTab% = 0) Then
                        Call EditOptionenLstMulti
                    ElseIf (hTab% = 4) Then
                        Call EditOptionenTxt
                    Else
                        Call EditOptionenLst
                    End If
                Case 2
                    If (hTab% = 3) And (UCase(s$) = "LIEFERANT") Then
                        Call EditOptionenLst
                    Else
                        Call EditOptionenTxt
                    End If
                Case 3
                    Call EditOptionenTxt
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

'If (Index < 3) Then
    If (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) Then
        cmdF5.Value = True
    End If
'End If
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

'If (Index < 3) Then
    If (KeyAscii = Asc("<")) Then
        cmdChange(0).Value = True
    ElseIf (KeyAscii = Asc(">")) Then
        cmdChange(1).Value = True
    End If
'End If

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
Dim h$, h2$, FormStr$


Call wpara.InitFont(Me)

Hoehe2% = 0
Breite2% = 0

tabOptionen.Left = wpara.LinksX
tabOptionen.Top = wpara.TitelY

For j% = 0 To 4
    With flxOptionen(j%)
        .Rows = 2
        .FixedRows = 1
        
        .Top = 3 * wpara.TitelY
        .Left = wpara.LinksX
        .Height = .RowHeight(0) * 11 + 90
        
        If (j% = 0) Then
            .Cols = 3
            FormStr$ = "^Tätigkeit|^Personal"
        ElseIf (j% = 1) Then
            .Cols = 3
            FormStr$ = "^Nr|^Preisbasis|^+ Aufschlag in %"
        ElseIf (j% = 2) Then
            .Cols = 5
            FormStr$ = "^Wert1|^Bedingung|^Wert2|^Rundung"
        ElseIf (j% = 3) Then
            .Cols = 3
            FormStr$ = "^Wert1|^Bedingung|^Wert2"
        Else
            .Cols = 2
            FormStr$ = "^Laufzeit bis (Monate)|^Warnung ab (Monate)"
        End If
        FormStr$ = FormStr$ + "|^ "
       
        .FormatString = FormStr$
        .Rows = 1
        .SelectionMode = flexSelectionFree
    End With
Next j%
        


Font.Bold = False   ' True

tabOptionen.Tab = 0
With flxOptionen(0)
    .Width = TextWidth(String(40, "W")) + 90    '(45
    .ColWidth(0) = 0
    .ColWidth(2) = 0
    For i% = 1 To (.Cols - 2)
        .ColWidth(i%) = .Width / 2
    Next i%

    spBreite% = 0
    For i% = 1 To .Cols - 2
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .ColWidth(0) = .Width - spBreite% - 90
    
    .Rows = 1
End With

tabOptionen.Tab = 1
With flxOptionen(1)
    .Width = flxOptionen(0).Width
    .ColWidth(0) = TextWidth("9999")
    .ColWidth(1) = .Width / 2
    .ColWidth(2) = 0
    .ColWidth(3) = 0

    spBreite% = 0
    For i% = 0 To .Cols - 2
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .ColWidth(2) = .Width - spBreite% - 90
End With

tabOptionen.Tab = 2
With flxOptionen(2)
    .Width = flxOptionen(0).Width
    
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth(" Bedingung  ")
    .ColWidth(2) = TextWidth(" Bedingung  ")
    .ColWidth(3) = TextWidth(" Bedingung       ")
    .ColWidth(4) = 0
    
'    For i% = 1 To (.Cols - 2)
'        .ColWidth(i%) = .Width / 3
'    Next i%

    spBreite% = 0
    For i% = 1 To .Cols - 2
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .ColWidth(0) = .Width - spBreite% - 90
End With

tabOptionen.Tab = 3
With flxOptionen(3)
    .Width = flxOptionen(0).Width
    
    For i% = 0 To 2
        .ColWidth(i%) = .Width / 3
    Next i%
End With

tabOptionen.Tab = 4
With flxOptionen(4)
    .Width = flxOptionen(0).Width
    
    For i% = 0 To 1
        .ColWidth(i%) = .Width / 2
    Next i%
End With


''''''''''''''''''''''''''''''''
tabOptionen.Tab = 5

chkSonstiges(0).Top = 3 * wpara.TitelY
chkSonstiges(0).Left = wpara.LinksX
For i% = 1 To 4
    chkSonstiges(i%).Top = chkSonstiges(i% - 1).Top + chkSonstiges(i% - 1).Height + 90
    chkSonstiges(i%).Left = chkSonstiges(i% - 1).Left
Next i%
chkSonstiges(5).Top = chkSonstiges(0).Top
chkSonstiges(5).Left = chkSonstiges(3).Left + chkSonstiges(3).Width - 180 '+ 180
For i% = 6 To 10
    chkSonstiges(i%).Top = chkSonstiges(i% - 1).Top + chkSonstiges(i% - 1).Height + 90
    chkSonstiges(i%).Left = chkSonstiges(i% - 1).Left
Next i%

txtSonstiges(0).Top = chkSonstiges(4).Top + chkSonstiges(4).Height + 180 ' 360
lblSonstiges(0).Left = wpara.LinksX
lblSonstiges(0).Top = txtSonstiges(0).Top + (txtSonstiges(0).Height - lblSonstiges(0).Height) / 2
txtSonstiges(0).Left = lblSonstiges(0).Left + lblSonstiges(0).Width + 150
lblSonstigesRechts(0).Left = txtSonstiges(0).Left + txtSonstiges(0).Width + 150
lblSonstigesRechts(0).Top = lblSonstiges(0).Top

txtSonstiges(1).Top = txtSonstiges(0).Top + txtSonstiges(0).Height + 90
lblSonstiges(1).Left = wpara.LinksX
lblSonstiges(1).Top = txtSonstiges(1).Top + (txtSonstiges(1).Height - lblSonstiges(1).Height) / 2
txtSonstiges(1).Left = txtSonstiges(0).Left     'lblSonstiges(1).Left + lblSonstiges(1).Width + 150
lblSonstigesRechts(1).Left = txtSonstiges(1).Left + txtSonstiges(1).Width + 150
lblSonstigesRechts(1).Top = lblSonstiges(1).Top



MaxWi% = 0
For i% = 2 To 3
    wi% = lblSonstiges(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

cboSonstiges(0).Top = txtSonstiges(1).Top + txtSonstiges(1).Height + 90
lblSonstiges(2).Left = wpara.LinksX
lblSonstiges(2).Top = cboSonstiges(0).Top + (cboSonstiges(0).Height - lblSonstiges(2).Height) / 2
cboSonstiges(0).Left = lblSonstiges(2).Left + MaxWi% + 150

cboSonstiges(1).Top = cboSonstiges(0).Top + cboSonstiges(0).Height + 90
lblSonstiges(3).Left = wpara.LinksX
lblSonstiges(3).Top = cboSonstiges(1).Top + (cboSonstiges(1).Height - lblSonstiges(3).Height) / 2
cboSonstiges(1).Left = cboSonstiges(0).Left

Hoehe2% = cboSonstiges(1).Top + cboSonstiges(1).Height + 180

'''''''''''''''''''''''''''''''''
   
Call OptionenBefuellen
'

fmeWuOptionen(0).Left = wpara.LinksX
fmeWuOptionen(0).Top = 900


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdF2.Width = TextWidth(cmdF5.Caption) + 150
cmdF2.Height = wpara.ButtonY
cmdF2.Left = tabOptionen.Left + fmeWuOptionen(0).Left + flxOptionen(0).Left + flxOptionen(0).Width + 150
cmdF2.Top = tabOptionen.Top + fmeWuOptionen(0).Top + flxOptionen(0).Top

Hoehe1% = flxOptionen(2).Top + flxOptionen(2).Height + 180
Breite1% = cmdF2.Left + cmdF2.Width + wpara.LinksX - tabOptionen.Left

Breite2% = fmeWuOptionen(0).Left + chkSonstiges(7).Left + chkSonstiges(7).Width + wpara.LinksX

'------------------------

tabOptionen.Tab = 3


''''''''''''''''''''''''''''


With fmeWuOptionen(0)
    .Left = wpara.LinksX
    .Top = 900
    
    If (Hoehe1% > Hoehe2%) Then
        .Height = Hoehe1%
    Else
        .Height = Hoehe2%
    End If
    If (Breite1% > Breite2%) Then
        .Width = Breite1%
    Else
        .Width = Breite2%
    End If
    
    .Caption = ""
End With
For i% = 1 To 5
    With fmeWuOptionen(i%)
        .Left = fmeWuOptionen(0).Left
        .Top = fmeWuOptionen(0).Top
        .Width = fmeWuOptionen(0).Width
        .Height = fmeWuOptionen(0).Height
        .Caption = ""
    End With
Next i%

With tabOptionen
    .Width = fmeWuOptionen(0).Left + fmeWuOptionen(0).Width + 2 * wpara.LinksX
    .Height = fmeWuOptionen(0).Top + fmeWuOptionen(0).Height + wpara.TitelY
End With


cmdF5.Width = cmdF2.Width
cmdF5.Height = cmdF2.Height
cmdF5.Left = cmdF2.Left
cmdF5.Top = cmdF2.Top + cmdF2.Height + 90

cmdAMPV.Width = cmdF2.Width
cmdAMPV.Height = cmdF2.Height
cmdAMPV.Left = cmdF2.Left
cmdAMPV.Top = cmdF2.Top

For i% = 0 To 1
    With cmdChange(i%)
        .Width = cmdF2.Width
        .Height = cmdF2.Height
        .Left = cmdF2.Left
    End With
Next i%
cmdChange(1).Top = tabOptionen.Top + fmeWuOptionen(0).Top + flxOptionen(0).Top + flxOptionen(0).Height - cmdChange(1).Height
cmdChange(0).Top = cmdChange(1).Top - cmdChange(0).Height - 90



'''''''''''''''''''''''''''''


'If (Hoehe1% > Hoehe2%) Then
'    tabOptionen.Height = Hoehe1%
'Else
'    tabOptionen.Height = Hoehe2%
'End If
'If (Breite1% > Breite2%) Then
'    tabOptionen.Width = Breite1%
'Else
'    tabOptionen.Width = Breite2%
'End If


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


If (OptionenModus% = 0) Then
    tabOptionen.Tab = 0
    Call TabDisable
    fmeWuOptionen(0).Visible = True
'    flxOptionen(0).Visible = True
    cmdF2.Visible = True
    cmdF5.Visible = True
Else
    tabOptionen.Tab = 1
    Call TabDisable
    fmeWuOptionen(1).Visible = True
'    flxOptionen(1).Visible = True
    cmdAMPV.Visible = True
    cmdF2.Visible = False
    cmdF5.Visible = True
    tabOptionen.TabEnabled(0) = False
    tabOptionen.TabEnabled(2) = False
    tabOptionen.TabEnabled(3) = False
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

If (tabOptionen.Tab = 5) Then chkSonstiges(0).SetFocus

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
Dim i%, j%, k%, l%, ind%
Dim h$, h2$, h3$

lstOptionen.Visible = False
lstOptionenMulti.Visible = False
txtOptionen.Visible = False

For j% = 0 To 4
    With flxOptionen(j%)
        .Rows = 1
        If (j% = 0) Then
            For i% = 0 To (AnzTaetigkeiten% - 1)
                h$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
                h$ = h$ + vbTab
                h2$ = ""
                For k% = 0 To 49
                    If (Taetigkeiten(i%).pers(k%) > 0) Then
                        h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
                    Else
                        Exit For
                    End If
                Next k%
                If (k% = 1) Then
                    ind% = Taetigkeiten(i%).pers(0)
                    h3$ = RTrim$(para.Personal(ind%))
                    h$ = h$ + h3$
                ElseIf (k% > 1) Then
                    h$ = h$ + "mehrere (" + Mid$(Str$(k%), 2) + ")"
                End If
                h$ = h$ + vbTab + h2$
                .AddItem h$
            Next i%
            .Rows = MAX_TAETIGKEITEN
        ElseIf (j% = 1) Then
            For i% = 0 To (MAX_AUFSCHLAEGE - 1)
                If (i% = (MAX_AUFSCHLAEGE - 1)) Then
                    h$ = "v"
                Else
                    h$ = Str$(i% + 1)
                End If
                ind% = AufschlagsTabelle(i%).PreisBasis
                If (ind% > 0) Then
                    If (ind% = 1) Then
                        h2$ = "NN-AEP"
                    ElseIf (ind% = 2) Then
                        h2$ = "Stamm-AEP"
                    Else
                        h2$ = "Taxe-AEP"
                    End If
                    h$ = h$ + vbTab + h2$ + vbTab
                    If (AufschlagsTabelle(i%).Aufschlag = 211) Then
                        h2$ = "AMPV"
                    Else
                        h2$ = Str$(AufschlagsTabelle(i%).Aufschlag)
                    End If
                    h$ = h$ + h2$
                End If
                .AddItem h$
            Next i%
            .Rows = MAX_AUFSCHLAEGE + 1
            .FillStyle = flexFillRepeat
            .row = .Rows - 1
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellBackColor = vbWhite
            .FillStyle = flexFillSingle
        ElseIf (j% = 2) Then
            For i% = 0 To (AnzRundungen% - 1)
                h$ = RTrim$(Rundungen(i%).bedingung.wert1)
                h$ = h$ + vbTab + RTrim$(Rundungen(i%).bedingung.op)
                h$ = h$ + vbTab + RTrim$(Rundungen(i%).bedingung.wert2)
                h$ = h$ + vbTab + RTrim$(Rundungen(i%).Gerundet)
                .AddItem h$
            Next i%
            .Rows = MAX_RUNDUNGEN
        ElseIf (j% = 3) Then
            For i% = 0 To (AnzWuSortierungen% - 1)
                h$ = RTrim$(WuSortierungen(i%).bedingung.wert1)
                h$ = h$ + vbTab + RTrim$(WuSortierungen(i%).bedingung.op)
                h$ = h$ + vbTab + RTrim$(WuSortierungen(i%).bedingung.wert2)
                flxOptionen(j%).AddItem h$
            Next i%
            flxOptionen(j%).Rows = MAX_WU_SORTIERUNGEN
        Else
            For i% = 0 To (AnzVerfallWarnungen% - 1)
                h$ = Format(VerfallWarnungen(i%).Laufzeit, "0")
                h$ = h$ + vbTab + Format(VerfallWarnungen(i%).Warnung, "0")
                flxOptionen(j%).AddItem h$
            Next i%
            flxOptionen(j%).Rows = MAX_VERFALL_WARNUNGEN
        End If
        .row = 1
        .col = 0
    End With
Next j%

chkSonstiges(0).Value = Abs(MacheEtiketten%)
chkSonstiges(1).Value = Abs(DruckeLagerKontrollListe%)
chkSonstiges(2).Value = Abs(BvProtAktiv%)
chkSonstiges(3).Value = Abs(TeilDefekte%)
chkSonstiges(4).Value = Abs(vAnzeigeSperren%)
chkSonstiges(5).Value = Abs(ModemInDOS%)
chkSonstiges(6).Value = Abs(Wbestk2ManuellSenden%)
chkSonstiges(7).Value = Abs(PartnerTeilBestellungen%)
chkSonstiges(8).Value = Abs(PartnerBestaendeBeruecksichtigen%)
chkSonstiges(9).Value = Abs(KalkNichtRezPflichtigeAM%)
chkSonstiges(10).Value = Abs(LieferantenAbfrage%)
txtSonstiges(0).text = Mid$(Str$(TageSpeichern%), 2)
txtSonstiges(1).text = Mid$(Str$(AnzRetourenDruck%), 2)

ind% = -1
For i% = 0 To (Printers.Count - 1)
    h$ = Printers(i%).DeviceName
    cboSonstiges(0).AddItem h$
    If (h$ = AutomatikDrucker$) Then
        cboSonstiges(0).ListIndex = i%
        ind% = i%
    End If
Next i%
If (ind% = -1) And (AutomatikDrucker$ <> "") Then
    cboSonstiges(0).AddItem AutomatikDrucker$, 0
    cboSonstiges(0).ListIndex = 0
End If

With cboSonstiges(1)
    .AddItem "Verfall"
    .AddItem "nächste LM"
    .AddItem "nicht springen"
    .ListIndex = NachManuellerLM%
End With

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
Dim h$, Send$, BetrLief$, Lief2$, BetrTage$, tag$, Aktiv$

With flxOptionen(0)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Taetigkeiten(j%).Taetigkeit = h$
            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 2)))
            For k% = 0 To 49
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
                    Taetigkeiten(j%).pers(k%) = Val(Lief2$)
                End If
            Next k%
            Do
                If (k% > 49) Then Exit Do
                Taetigkeiten(j%).pers(k%) = 0
                k% = k% + 1
            Loop
            j% = j% + 1
        End If
    Next i%
End With
AnzTaetigkeiten = j%

With flxOptionen(1)
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 1))
        If (h$ <> "") Then
            If (h$ = "NN-AEP") Then
                ind% = 1
            ElseIf (h$ = "Stamm-AEP") Then
                ind% = 2
            ElseIf (h$ = "Taxe-AEP") Then
                ind% = 3
            End If
        Else
            ind% = 0
        End If
        AufschlagsTabelle(i% - 1).PreisBasis = ind%
        If (ind% > 0) Then
            h$ = RTrim$(.TextMatrix(i%, 2))
            If (h$ = "AMPV") Then
                AufschlagsTabelle(i% - 1).Aufschlag = 211
            Else
                AufschlagsTabelle(i% - 1).Aufschlag = Val(h$)
            End If
        End If
    Next i%
End With

With flxOptionen(2)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Rundungen(j%).bedingung.wert1 = h$
            Rundungen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            Rundungen(j%).bedingung.wert2 = h$
            
            h$ = RTrim$(.TextMatrix(i%, 3))
            Rundungen(j%).Gerundet = h$
            
            j% = j% + 1
        End If
    Next i%
End With
AnzRundungen% = j%

With flxOptionen(3)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            WuSortierungen(j%).bedingung.wert1 = h$
            WuSortierungen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            WuSortierungen(j%).bedingung.wert2 = h$
            
            j% = j% + 1
        End If
    Next i%
End With
AnzWuSortierungen% = j%

With flxOptionen(4)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            VerfallWarnungen(j%).Laufzeit = Val(h$)
            h$ = RTrim$(.TextMatrix(i%, 1))
            VerfallWarnungen(j%).Warnung = Val(h$)
            
            j% = j% + 1
        End If
    Next i%
End With
AnzVerfallWarnungen% = j%

MacheEtiketten% = chkSonstiges(0).Value
DruckeLagerKontrollListe% = chkSonstiges(1).Value
BvProtAktiv% = chkSonstiges(2).Value
TeilDefekte% = chkSonstiges(3).Value
vAnzeigeSperren% = chkSonstiges(4).Value
ModemInDOS% = chkSonstiges(5).Value
Wbestk2ManuellSenden% = chkSonstiges(6).Value
PartnerTeilBestellungen% = chkSonstiges(7).Value
PartnerBestaendeBeruecksichtigen% = chkSonstiges(8).Value
KalkNichtRezPflichtigeAM% = chkSonstiges(9).Value
LieferantenAbfrage% = chkSonstiges(10).Value
TageSpeichern% = Val(txtSonstiges(0).text)
AnzRetourenDruck% = Val(txtSonstiges(1).text)
            
AutomatikDrucker$ = Trim(cboSonstiges(0).text)

NachManuellerLM% = cboSonstiges(1).ListIndex

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

For i% = 0 To 5
    fmeWuOptionen(i%).Visible = False
Next i%

'For i% = 0 To 4
'    flxOptionen(i%).Visible = False
'Next i%

cmdF2.Visible = False
cmdF5.Visible = False
cmdAMPV.Visible = False
For i% = 0 To 1
    cmdChange(i%).Visible = False
Next i%

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

fmeWuOptionen(hTab%).Visible = True

If (hTab% < 5) Then
    flxOptionen(hTab%).Visible = True
    flxOptionen(hTab%).SetFocus
    If (hTab% = 1) Then
        flxOptionen(hTab%).col = 1
        cmdAMPV.Visible = True
        cmdF5.Visible = True
    Else
        cmdF2.Visible = True
        cmdF5.Visible = True
        If (hTab% = 3) Then
            For i% = 0 To 1
                cmdChange(i%).Visible = True
            Next i%
        End If
    End If
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
            If (hTab% = 0) Then
                .AddItem "Alternativ-Artikel"
                .AddItem "Altlasten löschen"
                .AddItem "Fertigmachen"
                .AddItem "Hinzufügen"
                .AddItem "Lieferanten-Umsatz"
                .AddItem "Preiskalkulation"
                .AddItem "Preiskalk-Besorger"
                .AddItem "Preiskontrolle"
                .AddItem "RM-Kontrolle"
            ElseIf (hTab% = 2) Then
                .AddItem "wenn"
                .AddItem "sonst"
                .AddItem String$(50, "-")
                .AddItem "1.1x"
                .AddItem "1.xx"
                .AddItem String$(50, "-")
                .AddItem "Abbruch:billiger(%)"
                .AddItem "Abbruch:billiger(DM)"
                .AddItem "Abbruch:teurer(%)"
                .AddItem "Abbruch:teurer(DM)"
            ElseIf (hTab% = 3) Then
'                .AddItem "Absagen"
'                .AddItem "Anfragen"
                .AddItem "akzept.Angebote"
                .AddItem "Angebotshinweise"
                .AddItem "Besorger"
                .AddItem "Text-Besorger"
                .AddItem "BTM"
                .AddItem "Manuelle"
                .AddItem "Ladenhüter"
                .AddItem "Lagerartikel"
                .AddItem "def.Lagerartikel"
                .AddItem "Lagerartikel neg.LS"
                .AddItem "Schnelldreher"
                .AddItem "def.Schnelldreher"
                .AddItem "Interim"
                .AddItem "Interne Streichung"
                .AddItem String$(50, "-")
                .AddItem "BM"
                .AddItem "EK"
                .AddItem "VK"
                .AddItem "Zeilenwert"
                .AddItem "Lagerstand"
                .AddItem String$(50, "-")
                .AddItem "Hersteller"
                .AddItem "Lieferant"
                .AddItem "Lagercode"
                .AddItem "Lagercode 2st."
                .AddItem "Warengruppe"
            End If
            
        Case 1
            If (hTab% = 1) Then
                .AddItem "NN-AEP"
                .AddItem "Stamm-AEP"
                .AddItem "Taxe-AEP"
            ElseIf (hTab% = 2) Then
                .AddItem "<"
                .AddItem "<="
                .AddItem "="
                .AddItem "<>"
                .AddItem ">="
                .AddItem ">"
            ElseIf (hTab% = 3) Then
                h$ = ProgrammChar$
                ProgrammChar$ = "B"
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
    
                    Case Else
                        .AddItem "<"
                        .AddItem "<="
                        .AddItem "="
                        .AddItem "<>"
                        .AddItem ">="
                        .AddItem ">"
        '                        .AddItem "?-fach"
                End Select
                ProgrammChar$ = h$
            End If
            
    
        Case 2
            If (s$ = "Lieferant") Then
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
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
        .Left = tabOptionen.Left + fmeWuOptionen(hTab%).Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
'        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(1)
        .Top = tabOptionen.Top + fmeWuOptionen(hTab%).Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        If (xBreit%) Then
            .Width = TextWidth("Zeilenwert-Auftrag (ZW)") + wpara.FrmScrollHeight + 90
        Else
            .Width = flxOptionen(hTab%).ColWidth(col%)
        End If
'        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(1)
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowHeight(0)
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
        Case 1
            .AddItem "(keiner)"
            For i% = 1 To 50
                h$ = para.Personal(i%)
                .AddItem h$
            Next i%
    End Select

    For i% = 0 To (.ListCount - 1)
        .Selected(i%) = False
    Next i%

    
    Load frmEdit
    
     .ListIndex = 0
     
     BetrLief$ = LTrim$(RTrim$(flxOptionen(hTab%).TextMatrix(row%, 2)))
     
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
            .Selected(ind%) = True
         End If
     Next i%
     
    With frmEdit
        .Left = tabOptionen.Left + fmeWuOptionen(hTab%).Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
'        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(1)
        .Top = tabOptionen.Top + fmeWuOptionen(hTab%).Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen(hTab%).ColWidth(col%)
'        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(1)
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowHeight(0)
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
        If (EditGef%(0) = 255) Then
            flxOptionen(hTab%).TextMatrix(row%, col% + 1) = "255"
            h$ = "Naechstliefernder"
        Else
            If (EditAnzGefunden% = 0) Then
                h$ = ""
            ElseIf (EditAnzGefunden% = 1) Then
                ind% = EditGef%(0)
                h$ = RTrim$(para.Personal(ind%))
            Else
                h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
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

EditModus% = 4

hTab% = tabOptionen.Tab
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
        Case 0, 1
            .MaxLength = 3
            h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
            .text = h$
            EditModus% = 1
        Case 2
            If (hTab% = 1) Then
                .MaxLength = 0
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .text = h$
            ElseIf (hTab% = 2) Then
                h$ = Trim(flxOptionen(hTab%).TextMatrix(row%, col%))
                If (h$ = "") Then
                    If (iZeilenTyp% = 0) Then
                        .MaxLength = 1
                        h$ = "0"
                    ElseIf (iZeilenTyp% = 1) Then
                        .MaxLength = 2
                        h$ = "00"
                    Else
                        .MaxLength = 3
                        h$ = "000"
                    End If
                Else
                    .MaxLength = 0
                End If
                .text = h$
            ElseIf (hTab% = 3) Then
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
        Case 3
            If (iZeilenTyp% > 1) Then Call DefErrPop: Exit Sub

            .MaxLength = 4
            h$ = Trim(flxOptionen(hTab%).TextMatrix(row%, col%))
            If (h$ = "") Then
                If (iZeilenTyp% = 0) Then
                    h$ = "1.10"
                ElseIf (iZeilenTyp% = 1) Then
                    h$ = "1.00"
                End If
            Else
                .MaxLength = 0
            End If
            .text = h$
    End Select


    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + fmeWuOptionen(hTab%).Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + fmeWuOptionen(hTab%).Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
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

        h$ = UCase$(Trim$(EditTxt$))
        
'        If (col% = 1) Or ((col% = 2) And ((hTab% = 2) Or (iZeilenTyp% = 0))) Then
'            uhr% = Val(EditTxt$)
'            st% = uhr% \ 100
'            min% = uhr% Mod 100
'            h$ = Format(st%, "00") + ":" + Format(min%, "00")
'        ElseIf (iZeilenTyp% = 1) Then
'            h$ = UCase$(Trim$(EditTxt$))
'        Else
'            h$ = Mid$(Str(Val(EditTxt$)), 2)
'        End If
        With flxOptionen(hTab%)
            flxOptionen(hTab%).TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
    End If

End With

Call DefErrPop
End Sub

Private Sub txtsonstiges_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtSonstiges_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With txtSonstiges(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

