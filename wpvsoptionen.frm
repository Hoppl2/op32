VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWPvsOptionen 
   Caption         =   "Optionen"
   ClientHeight    =   5985
   ClientLeft      =   1770
   ClientTop       =   660
   ClientWidth     =   8565
   Icon            =   "wpvsoptionen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdNeu 
      Caption         =   "F2 neu"
      Height          =   450
      Left            =   5880
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdEntf 
      Caption         =   "F5 Löschen"
      Height          =   450
      Left            =   3960
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2280
      TabIndex        =   14
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   13
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   8085
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   14261
      _Version        =   327681
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
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
      TabPicture(0)   =   "wpvsoptionen.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeOptionen(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Überleitung"
      TabPicture(1)   =   "wpvsoptionen.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeOptionen(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Sonderprämie"
      TabPicture(2)   =   "wpvsoptionen.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeOptionen(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Personalfarben"
      TabPicture(3)   =   "wpvsoptionen.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeOptionen(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5 - Apotheke"
      TabPicture(4)   =   "wpvsoptionen.frx":037A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fmeOptionen(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame fmeOptionen 
         Height          =   4455
         Index           =   4
         Left            =   600
         TabIndex        =   32
         Top             =   840
         Width           =   8895
         Begin VB.TextBox txtOptionen4 
            Height          =   495
            Index           =   0
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "99"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtOptionen4 
            Height          =   495
            Index           =   1
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   36
            Text            =   "99"
            Top             =   960
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid flxOptionen4 
            Height          =   2700
            Left            =   1800
            TabIndex        =   37
            Top             =   1680
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
            _Version        =   65541
            Rows            =   7
            Cols            =   5
            BackColor       =   -2147483633
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            HighLight       =   2
            GridLines       =   0
            GridLinesFixed  =   1
            ScrollBars      =   0
         End
         Begin VB.Label lblOptionen4 
            Caption         =   "Toleranz für Öffnungs&zeiten (Minuten)"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   33
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblOptionen4 
            Caption         =   "Toleranz für &Vergleiche (%)"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   35
            Top             =   1080
            Width           =   2175
         End
      End
      Begin VB.Frame fmeOptionen 
         Caption         =   "&Sonderprämienartikel"
         Height          =   4815
         Index           =   2
         Left            =   -74640
         TabIndex        =   28
         Top             =   600
         Width           =   9855
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   3540
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   6244
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
         End
         Begin VB.Label lblOptionen3 
            Caption         =   "ACHTUNG: die Veränderung der Sonderprämienartikel gilt nicht für bereits übergeleitete Daten !!!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   7575
         End
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   3
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   8895
         Begin VB.ComboBox cboOptionen1 
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown-Liste
            TabIndex        =   25
            Top             =   1440
            Width           =   2415
         End
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   2
            Left            =   720
            TabIndex        =   26
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
         Begin MSComDlg.CommonDialog CommDlg 
            Left            =   4560
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   327681
         End
         Begin VB.Label lblOptionen1 
            Caption         =   "&Position der Legenden"
            Height          =   375
            Index           =   0
            Left            =   3360
            TabIndex        =   27
            Top             =   1440
            Width           =   1815
         End
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4815
         Index           =   1
         Left            =   -74760
         TabIndex        =   17
         Top             =   600
         Width           =   9855
         Begin VB.CheckBox chkOptionen 
            Caption         =   "&Privatrezepte"
            Height          =   495
            Index           =   2
            Left            =   6480
            TabIndex        =   20
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkOptionen 
            Caption         =   "&Rabattabzug"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkOptionen 
            Caption         =   "&Lieferscheine"
            Height          =   495
            Index           =   1
            Left            =   3960
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   3540
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   6244
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
         End
         Begin VB.Label lblOptionen 
            Caption         =   "ACHTUNG: die Veränderung dieser Parameter gilt nicht für bereits übergeleitete Daten !!!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   7575
         End
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4455
         Index           =   0
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   4
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   2
            Text            =   "WWW9999"
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   3
            Left            =   1680
            MaxLength       =   7
            TabIndex        =   0
            Text            =   "WWW9999"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   7
            Left            =   5760
            TabIndex        =   8
            Text            =   "WWW99,99"
            Top             =   1320
            Width           =   975
         End
         Begin VB.Frame fraBasis 
            Caption         =   "&Prämienbasis"
            Height          =   1575
            Left            =   4800
            TabIndex        =   9
            Top             =   2040
            Width           =   2535
            Begin VB.OptionButton optBasis 
               Caption         =   "&Spanne"
               Height          =   495
               Index           =   2
               Left            =   120
               TabIndex        =   12
               Top             =   960
               Width           =   1935
            End
            Begin VB.OptionButton optBasis 
               Caption         =   "AVP &exkl. MwSt."
               Height          =   495
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   600
               Width           =   1815
            End
            Begin VB.OptionButton optBasis 
               Caption         =   "AVP &inkl. MwSt."
               Height          =   495
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   5
            Left            =   5640
            MaxLength       =   38
            TabIndex        =   4
            Text            =   "WWW99,99"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   6
            Left            =   5880
            TabIndex        =   6
            Text            =   "WWW99,99"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&große Pause ab"
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   1
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Sonderprämie in %"
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   7
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Grundprämie in %"
            Height          =   375
            Index           =   5
            Left            =   4200
            TabIndex        =   3
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Zusatzprämie in %"
            Height          =   375
            Index           =   6
            Left            =   4320
            TabIndex        =   5
            Top             =   720
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmWPvsOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgRezApoDruckName$


Private Const DefErrModul = "WINPVSOPTIONEN.FRM"

Const flxPersonalInd% = 2
Const fmePersonal% = 3
Sub EditParams()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditParams")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, hTab%, Row%, Col%, aRow%, uhr%, st%, min%, iZeilenTyp%
Dim s$, h$

Row% = flxOptionen1(0).Row
Col% = flxOptionen1(0).Col
                
With frmEdit.txtEdit
    .MaxLength = 0
    h$ = Trim(flxOptionen1(0).TextMatrix(Row%, Col%))
    .Text = h$
    EditModus% = 1
    If Col% = 2 Then
        EditModus% = 6
        .MaxLength = 4
    ElseIf Col% >= 4 And Col% <= 8 Then
        EditModus% = 4
        .MaxLength = 9
    ElseIf Col% = 9 Then
      EditModus% = 3
      .MaxLength = 1
    End If
    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen1(0).Left + flxOptionen1(0).ColPos(Col%) + 45
        .Left = .Left + fmeOptionen(1).Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen1(0).Top + (Row% - flxOptionen1(0).TopRow + 1) * flxOptionen1(0).RowHeight(0)
        .Top = .Top + fmeOptionen(1).Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen1(0).ColWidth(Col%)
        If Col% = 0 Then
          .Height = 3 * flxOptionen1(0).RowHeight(0) + 100
        Else
          .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
        End If
    End With
    If Col% = 0 Then
      With frmEdit.lstEdit
        .Clear
        If Row% = 1 Then
          .AddItem "NICHT"
        Else
          .AddItem "UND"
          .AddItem "ODER"
          .AddItem "NICHT"
          .ListIndex = 0
          If h$ = "ODER" Then
            .ListIndex = 1
          ElseIf h$ = "NICHT" Then
            .ListIndex = 2
          End If
        End If
        .Width = frmEdit.ScaleWidth
        .Height = frmEdit.ScaleHeight
        .Left = 0
        .Top = 0
        .BackColor = vbWhite
        .Visible = True
      End With
    Else
      With frmEdit.txtEdit
          .Width = frmEdit.ScaleWidth
  '            .Height = frmEdit.ScaleHeight
          .Left = 0
          .Top = 0
          .BackColor = vbWhite
          .Visible = True
       End With
    End If
    frmEdit.Show 1
    
    If (EditErg%) Then
        h$ = Trim$(EditTxt$)
        With flxOptionen1(0)
            .TextMatrix(Row%, Col%) = h$
        End With
    End If

End With

Call DefErrPop

End Sub

Sub EinlesenSprmArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenSprmArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim erg%
Dim satz As Long

If sp.RecordCount > 0 Then
  sp.MoveFirst
  With flxOptionen1(1)
    Do While Not sp.EOF
      h$ = sp!pzn
      erg% = ast.IndexSearch(0, h$, satz)
      If erg% = 0 Then
        ast.GetRecord (satz + 1)
        h$ = h$ + vbTab + ast.kurz + vbTab + ast.meng + vbTab + ast.meh
        .AddItem h$
      End If
      sp.MoveNext
    Loop
    .Row = 1
    .RowSel = 0
    .Col = 1
    .ColSel = 3
    .Sort = 5
  End With
End If
Call DefErrPop
End Sub

Private Sub cmdEntf_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEntf_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim pzn As String

If tabOptionen.Tab = 1 Then
  With flxOptionen1(0)
    For i% = 0 To .Cols - 1
      .TextMatrix(.Row, i%) = ""
    Next i%
    .SetFocus
  End With
ElseIf tabOptionen.Tab = 2 Then
  With flxOptionen1(1)
    pzn = .TextMatrix(.Row, 0)
    sp.Index = "Unique"
    sp.Seek "=", pzn
    If Not sp.NoMatch Then
      sp.Delete
    End If
    If .Rows > 1 Then
      If .Rows = 2 Then   'letzte Zeile außer Fixedrow kann nicht gelöscht werden
        .Rows = 1
      Else
        .RemoveItem (.Row)
      End If
    End If
  End With
End If
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

Private Sub cmdNeu_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdNeu_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
frmSPrmArtikel.Show vbModal
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
Dim i%, j%, VERBAND%
Dim l&
Dim h$, s$, pzn As String, txt As String

If (ActiveControl.Name = cmdOk.Name) Then
    Call AuslesenFlexOptionen
    
    Call SpeicherIniPersonalFarben
    Call SpeicherIniLegendenPos
    Call SpeicherParameter
    Unload Me
ElseIf (ActiveControl.Name = flxOptionen1(flxPersonalInd%).Name) Then
    If (ActiveControl.Index = flxPersonalInd%) Then
        If (ActiveControl.Col = 0) Then
            Call EditPersonalFarbe
        Else
            Call EditInitialen
        End If
    ElseIf (ActiveControl.Index = 0) Then
        Call EditParams
    End If
'
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

With flxOptionen1(Index)
    If (Index = flxPersonalInd%) And (.Visible) And (.Enabled) Then
        If (.Col > 1) Then
            .Col = 1
            .ColSel = .Col
        End If
    End If
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxOptionen1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (Index = 0 Or Index = 1) And KeyCode = vbKeyF5 Then
  Call cmdEntf_Click
End If
If Index = 1 And KeyCode = vbKeyF2 Then
  Call cmdNeu_Click
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

With flxOptionen1(Index)
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub flxOptionen1_RowColChange(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_RowColChange")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxOptionen1(Index)
    If (Index = flxPersonalInd%) And (.Visible) And (.Enabled) Then
        If (.Col > 1) Then
            .Col = 1
            .ColSel = .Col
        End If
    End If
    .HighLight = flexHighlightAlways
End With

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
Dim w As Long
Dim h$, h2$, h3$, FormStr$
Dim c As Object


Call wpara.InitFont(Me)

'txtOptionen0(1).Text = String(38, "A")

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.Text) + 90
        c.Text = ""
    End If
Next
On Error GoTo DefErr

'''''''''''''''''''''''''''''''''''
With fmeOptionen(0)
    .Left = wpara.LinksX
    .Top = 3 * wpara.TitelY
'    .Width = txtOptionen0(1).Left + txtOptionen0(1).Width + 900
    .Width = flxOptionen1(0).Left + flxOptionen1(0).Width + 900
    .Height = fraBasis.Top + fraBasis.Height + 300
End With
For i% = 1 To 4
    With fmeOptionen(i%)
        .Left = fmeOptionen(0).Left
        .Top = fmeOptionen(0).Top
        .Width = fmeOptionen(0).Width
        .Height = fmeOptionen(0).Height
    End With
Next i%

tabOptionen.Tab = 0

lblOptionen0(0).Left = wpara.LinksX
lblOptionen0(0).Top = 2 * wpara.TitelY
txtOptionen0(0).Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 300
txtOptionen0(0).Top = lblOptionen0(0).Top + (lblOptionen0(0).Height - txtOptionen0(0).Height) / 2

For i% = 1 To 4
    lblOptionen0(i%).Left = lblOptionen0(0).Left
    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
    txtOptionen0(i%).Left = txtOptionen0(0).Left
    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
Next i%

lblOptionen0(5).Left = txtOptionen0(0).Left + txtOptionen0(0).Width + 300
lblOptionen0(5).Top = lblOptionen0(0).Top
txtOptionen0(5).Left = lblOptionen0(5).Left + lblOptionen0(5).Width + 150
txtOptionen0(5).Top = lblOptionen0(5).Top + (lblOptionen0(5).Height - txtOptionen0(5).Height) / 2
For i% = 6 To 7
    lblOptionen0(i%).Left = txtOptionen0(0).Left + txtOptionen0(0).Width + 300
    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
    txtOptionen0(i%).Left = txtOptionen0(5).Left
    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
Next i%

txtOptionen0(0).Text = Format(TagAnf%, "#0:00")
txtOptionen0(1).Text = Format(TagEnd%, "#0:00")
txtOptionen0(2).Text = Format(NORMTAGZEIT%, "####0")
txtOptionen0(3).Text = Format(PauseKl%, "##0")
txtOptionen0(4).Text = Format(PauseGr%, "##0")
txtOptionen0(5).Text = Format(pb#, "##0.00")
txtOptionen0(6).Text = Format(zpb#, "##0.00")
txtOptionen0(7).Text = Format(spb#, "##0.00")

fraBasis.Left = txtOptionen0(5).Left + txtOptionen0(5).Width + 300
fraBasis.Top = lblOptionen0(5).Top - 150
fraBasis.Width = lblOptionen0(5).Width + 150 + txtOptionen0(5).Width

optBasis(0).Top = 2 * wpara.TitelY
optBasis(0).Left = wpara.LinksX
optBasis(1).Top = optBasis(0).Top + optBasis(0).Height + 200
optBasis(1).Left = wpara.LinksX
optBasis(2).Top = optBasis(1).Top + optBasis(1).Height + 200
optBasis(2).Left = wpara.LinksX
fraBasis.Height = optBasis(2).Top + optBasis(2).Height + 200

Select Case PrämBasis$
Case "I"
  optBasis(0).Value = True
Case "E"
  optBasis(1).Value = True
Case "S"
  optBasis(2).Value = True
End Select

tabOptionen.Tab = 1
lblOptionen.Top = wpara.TitelY
lblOptionen.Left = wpara.LinksX

For i% = 0 To 2
  chkOptionen(i%).Top = lblOptionen.Top + lblOptionen.Height + wpara.TitelY
Next i%
chkOptionen(0).Left = wpara.LinksX

Me.Font.Bold = True

With flxOptionen1(0)
  .Rows = 2
  .Cols = 11
  .FixedRows = 1
  .FormatString = "^|<Warengruppe|<Tara|<Lagercodes|<Geräte|>AVP von|>AVP bis|>Spanne von|>Spanne bis|>RP|"
  .Rows = 11
  .Left = wpara.LinksX
  .Top = chkOptionen(0).Top + chkOptionen(0).Height + wpara.TitelY
  .Height = 11 * .RowHeight(0) + 100
  .SelectionMode = flexSelectionFree
  .ColWidth(0) = Me.TextWidth(" WWWWW ")
  .ColWidth(1) = Me.TextWidth(" WWWWWWW ")
  .ColWidth(2) = Me.TextWidth(" ABCD ")
  .ColWidth(3) = Me.TextWidth(" WWWWWW ")
  .ColWidth(4) = Me.TextWidth(" 123456 ")
  .ColWidth(5) = Me.TextWidth(" 9999999 ")
  .ColWidth(6) = Me.TextWidth(" 9999999 ")
  .ColWidth(7) = Me.TextWidth(" 9999999999 ")
  .ColWidth(8) = Me.TextWidth(" 9999999999 ")
  .ColWidth(9) = Me.TextWidth(" WWW ")
  .ColWidth(10) = wpara.FrmScrollHeight
  For i% = 0 To 10
    w = w + .ColWidth(i%)
  Next i%
  .Width = w + 50
  For i% = 1 To 10
    Select Case operator$(i%)
    Case "U"
      .TextMatrix(i%, 0) = " UND"
    Case "O"
      .TextMatrix(i%, 0) = " ODER"
    Case "N"
      .TextMatrix(i%, 0) = " NICHT"
    End Select
    .TextMatrix(i%, 1) = wg$(i%)
    .TextMatrix(i%, 2) = tLager$(i%)
    .TextMatrix(i%, 3) = LgCode$(i%)
    .TextMatrix(i%, 4) = Geräte$(i%)
    If vonAVP#(i%) > 0# Then .TextMatrix(i%, 5) = Format(vonAVP#(i%), "#######0")
    If bisAVP#(i%) < 9999999# Then .TextMatrix(i%, 6) = Format(bisAVP#(i%), "#######0")
    If vonSP#(i%) > 0# Then .TextMatrix(i%, 7) = Format(vonSP#(i%), "#######0")
    If bisSP#(i%) < 9999999# Then .TextMatrix(i%, 8) = Format(bisSP#(i%), "#######0")
    .TextMatrix(i%, 9) = Rp$(i%) + " "
  Next i%
  
  fmeOptionen(1).Width = .Left + .Width + 300
  If fmeOptionen(1).Width > fmeOptionen(0).Width Then
    fmeOptionen(0).Width = fmeOptionen(1).Width
    fmeOptionen(2).Width = fmeOptionen(1).Width
    fmeOptionen(3).Width = fmeOptionen(1).Width
    fmeOptionen(4).Width = fmeOptionen(1).Width
  End If
  .Height = fmeOptionen(1).Height - .Top - 150
End With
Me.Font.Bold = False

chkOptionen(1).Left = chkOptionen(0).Left + chkOptionen(0).Width
chkOptionen(1).Left = chkOptionen(1).Left + (chkOptionen(2).Left - chkOptionen(1).Left) \ 2 - (chkOptionen(1).Width \ 2)
chkOptionen(2).Left = flxOptionen1(0).Left + flxOptionen1(0).Width - wpara.LinksX - chkOptionen(2).Width
If RabAbzug Then chkOptionen(0).Value = 1
If LSAuch Then chkOptionen(1).Value = 1
If PrivRez Then chkOptionen(2).Value = 1

tabOptionen.Tab = 2
lblOptionen3.Top = wpara.TitelY + 150
lblOptionen3.Left = wpara.LinksX

With flxOptionen1(1)
  .Top = lblOptionen3.Top + lblOptionen3.Height + 150
  .Rows = 2
  .Cols = 5
  .FixedRows = 1
  .FormatString = "<PZN|<Bezeichnung|>Mng|<EH|"
  .Rows = 1
  
  .ColWidth(0) = TextWidth("999999999")
  .ColWidth(1) = 0
  .ColWidth(2) = TextWidth("WWWWWW")
  .ColWidth(3) = TextWidth("WWW")
  .ColWidth(4) = wpara.FrmScrollHeight

  wi% = 90
  For i% = 0 To (.Cols - 1)
      wi% = wi% + .ColWidth(i%)
  Next i%
  .Width = fmeOptionen(2).Width - 2 * wpara.LinksX
  .ColWidth(1) = .Width - wi%
  .Height = fmeOptionen(2).Height - .Top - 150
End With
Call EinlesenSprmArtikel

tabOptionen.Tab = 3
With flxOptionen1(flxPersonalInd%)
    .Rows = 2
    .Cols = 6
    .FixedRows = 1
    .FormatString = "<PersonalName|^Initialen|^Farbe am PC||"
    .Rows = 1
    
    For i% = 1 To 50
        h$ = Trim(para.Personal(i%))
        If (h$ <> "") Then
            .AddItem para.Personal(i%) + vbTab + PersonalInitialen$(i%) + vbTab + vbTab + Str$(i%) + vbTab + PersonalFarben$(i%)
            .FillStyle = flexFillRepeat
            .Row = .Rows - 1
            .Col = 2
            .RowSel = .Row
            .ColSel = .Cols - 1
            .CellBackColor = wpara.BerechneFarbWert(PersonalFarben$(i%))
            .FillStyle = flexFillSingle
        End If
    Next i%
    
    Breite1% = 0
    For i% = 0 To (.Rows - 1)
        Breite2% = TextWidth(.TextMatrix(i%, 0))
        If (Breite2% > Breite1%) Then Breite1% = Breite2%
    Next i%
    .ColWidth(0) = Breite1% + 450
    .ColWidth(1) = TextWidth(String(5, "W"))
    .ColWidth(2) = .ColWidth(0) ' TextWidth(String(18, "X"))
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    .ColWidth(5) = wpara.FrmScrollHeight

    wi% = 90
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi%
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 90
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    .Height = fmeOptionen(3).Height - .Top - 150
    
    .SelectionMode = flexSelectionFree
    .Row = 1
    .Col = 0
    .ColSel = .Col
    
End With

With lblOptionen1(0)
    .Left = flxOptionen1(flxPersonalInd%).Left + flxOptionen1(flxPersonalInd%).Width + 300
    .Top = flxOptionen1(flxPersonalInd%).Top
End With
With cboOptionen1
    .Left = lblOptionen1(0).Left
    .Top = lblOptionen1(0).Top + lblOptionen1(0).Height + 300
    .AddItem "Links"
    .AddItem "Rechts"
    .AddItem "Oben"
    .AddItem "Unten"
    .AddItem "Ausblenden"
    .ListIndex = 0
    For i% = 1 To (.ListCount - 1)
        If (Left$(.List(i%), 1) = LegendenPosStr$) Then
            .ListIndex = i%
            Exit For
        End If
    Next i%
End With



tabOptionen.Tab = 4

lblOptionen4(0).Left = wpara.LinksX
lblOptionen4(0).Top = 2 * wpara.TitelY
txtOptionen4(0).Left = lblOptionen4(0).Left + lblOptionen4(0).Width + 300
txtOptionen4(0).Top = lblOptionen4(0).Top + (lblOptionen4(0).Height - txtOptionen4(0).Height) / 2

For i% = 1 To 1
    lblOptionen4(i%).Left = lblOptionen4(0).Left
    lblOptionen4(i%).Top = lblOptionen4(i% - 1).Top + lblOptionen4(i% - 1).Height + 300
    txtOptionen4(i%).Left = txtOptionen4(0).Left
    txtOptionen4(i%).Top = lblOptionen4(i%).Top + (lblOptionen4(i%).Height - txtOptionen4(i%).Height) / 2
Next i%

With flxOptionen4
    .Rows = 7
    .Cols = 5
    .FixedRows = 1
    .FixedCols = 1
    .FormatString = "<Öffnungszeiten|^von|^bis|^von|^bis"
    
    .SelectionMode = flexSelectionFree
    .Row = 1
    .Col = 1
    .ColSel = .Col
End With

'''''''''''''''''''''''''''''''''''

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)


With tabOptionen
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = fmeOptionen(0).Left + fmeOptionen(0).Width + wpara.LinksX
    .Height = fmeOptionen(0).Top + fmeOptionen(0).Height + wpara.TitelY
End With

cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
cmdEsc.Top = cmdOk.Top
cmdEntf.Top = cmdOk.Top
cmdNeu.Top = cmdOk.Top

Me.Width = tabOptionen.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEntf.Width = wpara.ButtonX
cmdEntf.Height = wpara.ButtonY
cmdNeu.Width = wpara.ButtonX
cmdNeu.Height = wpara.ButtonY

cmdEsc.Left = tabOptionen.Left + tabOptionen.Width - cmdEsc.Width
cmdOk.Left = cmdEsc.Left - cmdOk.Width - 150
'cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
'cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
cmdNeu.Left = tabOptionen.Left
cmdEntf.Left = cmdNeu.Left + cmdNeu.Width + 150

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
If (Breite1% < 0) Then Breite1% = 0
Me.Left = Breite1%
Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
If (Hoehe1% < 0) Then Hoehe1% = 0
Me.Top = Hoehe1%

tabOptionen.Tab = 0
Call TabDisable
Call TabEnable(tabOptionen.Tab)

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

Call DefErrPop
End Sub

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

For i% = 0 To 4
    fmeOptionen(i%).Visible = False
Next i%
cmdEntf.Visible = False
cmdNeu.Visible = False
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

fmeOptionen(hTab%).Visible = True

If (hTab% = 0) Then
    If (txtOptionen0(0).Visible) Then txtOptionen0(0).SetFocus
ElseIf (hTab% = 1) Then
    flxOptionen1(0).SetFocus
    cmdEntf.Visible = True
ElseIf (hTab% = 2) Then
    flxOptionen1(1).SetFocus
    cmdEntf.Visible = True
    cmdNeu.Visible = True
ElseIf (hTab% = 3) Then
    flxOptionen1(2).SetFocus
ElseIf (hTab% = 4) Then
    txtOptionen4(0).SetFocus
End If

Call DefErrPop
End Sub

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

With txtOptionen0(Index)
    h$ = .Text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .Text = h$
    .SelStart = 0
    .SelLength = Len(.Text)
End With

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

If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((Index = 0) Or (Chr$(KeyAscii) <> ".")) And ((Index > 0) Or (Chr$(KeyAscii) <> ":")) Then
    Beep
    KeyAscii = 0
End If

Call DefErrPop
End Sub

Sub AuslesenFlexTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexTaetigkeiten")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim h$, BetrLief$, Lief2$, Key$

'With flxOptionen1(1)
'    j% = 0
'    For i% = 1 To (.Rows - 1)
'        h$ = RTrim$(.TextMatrix(i%, 0))
'        If (h$ <> "") And (RTrim$(.TextMatrix(i%, 1)) <> "") Then
'            Taetigkeiten(j%).Taetigkeit = h$
'            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 2)))
'            For k% = 0 To 49
'                If (BetrLief$ = "") Then Exit For
'
'                ind% = InStr(BetrLief$, ",")
'                If (ind% > 0) Then
'                    Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
'                    BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
'                Else
'                    Lief2$ = BetrLief$
'                    BetrLief$ = ""
'                End If
'                If (Lief2$ <> "") Then
'                    Taetigkeiten(j%).pers(k%) = Val(Lief2$)
'                End If
'            Next k%
'            Do
'                If (k% > 49) Then Exit Do
'                Taetigkeiten(j%).pers(k%) = 0
'                k% = k% + 1
'            Loop
'            j% = j% + 1
'        End If
'    Next i%
'End With
'AnzTaetigkeiten = j%
'
'i% = 1
'If (i% <= AnzTaetigkeiten) Then
'    h$ = RTrim$(Taetigkeiten(i% - 1).Taetigkeit)
'    For j% = 0 To 49
'        If (Taetigkeiten(i% - 1).pers(j%) > 0) Then
'            h$ = h$ + "," + Mid$(Str$(Taetigkeiten(i% - 1).pers(j%)), 2)
'        Else
'            Exit For
'        End If
'    Next j%
'Else
'    h$ = ""
'End If
'
'Key$ = "Taetigkeit" + Format(i%, "00")
'l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)

Call DefErrPop
End Sub

Sub EditPersonalFarbe()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditPersonalFarbe")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iPersonal%
Dim iFarbePersonal&
Dim l&

With flxOptionen1(flxPersonalInd%)
'    iPersonal% = Val(.TextMatrix(.row, 3))
    iFarbePersonal& = wpara.BerechneFarbWert(.TextMatrix(.Row, 4))
    On Error Resume Next
    CommDlg.Color = iFarbePersonal&
    CommDlg.CancelError = True
    CommDlg.Flags = cdlCCFullOpen + cdlCCRGBInit
    Call CommDlg.ShowColor
    If (Err = 0) Then
        iFarbePersonal& = CommDlg.Color
        .TextMatrix(.Row, 4) = Right$("000000" + Hex$(iFarbePersonal&), 6)
        .Enabled = False
        .FillStyle = flexFillRepeat
        .Col = 2
        .RowSel = .Row
        .ColSel = .Cols - 1
        .CellBackColor = iFarbePersonal&    ' wpara.BerechneFarbWert(PersonalFarben$(iPersonal%))
        .FillStyle = flexFillSingle
        .Enabled = True
        .Col = 0
        .SetFocus
    End If
End With


Call DefErrPop
End Sub

Sub EditInitialen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditInitialen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, hTab%, Row%, Col%, aRow%, uhr%, st%, min%, iZeilenTyp%
Dim s$, h$

Row% = flxOptionen1(flxPersonalInd%).Row
Col% = flxOptionen1(flxPersonalInd%).Col
                
With frmEdit.txtEdit
    .MaxLength = 3
    h$ = flxOptionen1(flxPersonalInd%).TextMatrix(Row%, Col%)
    .Text = Right$(Space$(3) + h$, 3)
    EditModus% = 1
    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen1(flxPersonalInd%).Left + flxOptionen1(flxPersonalInd%).ColPos(Col%) + 45
        .Left = .Left + fmeOptionen(fmePersonal%).Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen1(flxPersonalInd%).Top + (Row% - flxOptionen1(flxPersonalInd%).TopRow + 1) * flxOptionen1(flxPersonalInd%).RowHeight(0)
        .Top = .Top + fmeOptionen(fmePersonal%).Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen1(flxPersonalInd%).ColWidth(Col%)
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
    
    If (EditErg%) Then
        h$ = Trim$(EditTxt$)
        With flxOptionen1(flxPersonalInd%)
            .TextMatrix(Row%, Col%) = h$
        End With
    End If

End With

Call DefErrPop
End Sub


Sub AuslesenFlexOptionen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexOptionen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, iPersonal%
    
LegendenPosStr$ = Left$(cboOptionen1.Text, 1)

With flxOptionen1(flxPersonalInd%)
    For i% = 1 To (.Rows - 1)
        iPersonal% = Val(.TextMatrix(i%, 3))
        PersonalFarben$(iPersonal%) = .TextMatrix(i%, 4)
        PersonalInitialen$(iPersonal%) = Trim(.TextMatrix(i%, 1))
    Next i%
End With

Call DefErrPop
End Sub

