VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmEasyMatch 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6120
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF3 
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3240
      Picture         =   "EasyMatch.frx":0000
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
      Index           =   1
      Left            =   3480
      Picture         =   "EasyMatch.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3720
      Picture         =   "EasyMatch.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdF3 
      Caption         =   "&Edit (F3)"
      Height          =   450
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "&Löschen (F5)"
      Height          =   450
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "&Neu (F2)"
      Height          =   450
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxEasyMatch 
      Height          =   2700
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmEasyMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "EASYMATCH.FRM"

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

FormErg% = True
Call SpeicherEasyMatch
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
Dim iRow%

iRow% = 0
If (EasyMatchModus% = 2) Then
    iRow% = 1
End If

With flxEasyMatch
    If (.row > iRow%) Then
        .AddItem "", .row
        Call EditZeile
    End If
    .SetFocus
End With

Call DefErrPop
End Sub

Private Sub cmdF3_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF3_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iRow%

iRow% = 0
If (EasyMatchModus% = 2) Then
    iRow% = 1
End If

With flxEasyMatch
    If (.row > iRow%) Then
        Call EditZeile
    End If
    .SetFocus
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
Dim iRow%

iRow% = 0
If (EasyMatchModus% = 2) Then
    iRow% = 1
End If

With flxEasyMatch
    If (.row > iRow%) Then
        .RemoveItem .row
    End If
    .SetFocus
End With

Call DefErrPop
End Sub

Private Sub flxEasyMatch_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEasyMatch_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxEasyMatch
    .HighLight = flexHighlightAlways
    .col = 0
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub

Private Sub flxEasyMatch_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEasyMatch_LostFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxEasyMatch
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub flxEasyMatch_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEasyMatch_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, row%, gef%, col%
Dim ch$, h$

ch$ = UCase$(Chr$(KeyAscii))

If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxEasyMatch
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                    .row = i%
                    gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub flxEasyMatch_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEasyMatch_RowColChange")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxEasyMatch
'    If (.Redraw) And (.Visible) Then
'        .col = 0
'        .ColSel = .Cols - 1
'    End If
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%, iRow%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control

iEditModus = 1

FormErg% = False

Call wpara.InitFont(Me)

With flxEasyMatch
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Name|<Wert|"
    .Rows = 1
    
    If (EasyMatchModus% = 0) Then
        Caption = "Editierung Hilfsmittel-Nummern"
    ElseIf (EasyMatchModus% = 1) Then
        Caption = "Editierung Sonder-Pzn"
    Else
        Caption = "Editierung Kassen-Aufschläge"
    End If
    
    Call ActProgram.FlxEasyMatchBefuellen
    
    
    .ColWidth(0) = TextWidth(String(40, "A"))
    .ColWidth(1) = TextWidth(String(12, "9"))
    .ColWidth(2) = wpara.FrmScrollHeight
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    
    If (.Rows <= 10) Then
        .Height = .RowHeight(0) * 10 + 90
    ElseIf (.Rows <= 15) Then
        .Height = .RowHeight(0) * .Rows + 90
    Else
        .Height = .RowHeight(0) * 15 + 90
    End If
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    

    iRow% = 0
    If (EasyMatchModus% = 2) Then
        iRow% = 1
    End If

    .row = iRow% + 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    .col = 1
            
    .row = 1
End With

Font.Bold = False   ' True

With cmdF2
    .Left = flxEasyMatch.Left + flxEasyMatch.Width + 300
    .Top = flxEasyMatch.Top
    .Width = TextWidth(cmdF5.Caption) + 150
    .Height = wpara.ButtonY
End With

With cmdF3
    .Left = cmdF2.Left
    .Top = cmdF2.Top + cmdF2.Height + 150
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With

With cmdF5
    .Left = cmdF2.Left
    .Top = cmdF3.Top + cmdF3.Height + 150
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With


Me.Width = cmdF2.Left + cmdF2.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height

cmdOk.Top = flxEasyMatch.Top + flxEasyMatch.Height + 150
cmdEsc.Top = cmdOk.Top

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxEasyMatch
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    
    cmdF2.Left = cmdF2.Left + 2 * iAdd
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    
    Height = Height + iAdd2
    
    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxEasyMatch.Top + flxEasyMatch.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdOk
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdEsc.Top
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdF2
        .Init
        .Left = cmdF2.Left
        .Top = cmdF2.Top
        .Caption = cmdF2.Caption
        .TabIndex = cmdF2.TabIndex
        .Enabled = cmdF2.Enabled
        .Default = cmdF2.Default
        .Cancel = cmdF2.Cancel
        .Visible = True
    End With
    cmdF2.Visible = False

    With nlcmdF3
        .Init
        .Left = nlcmdF2.Left
        .Top = nlcmdF2.Top + nlcmdF2.Height + 90
        .Caption = cmdF3.Caption
        .TabIndex = cmdF3.TabIndex
        .Enabled = cmdF3.Enabled
        .Default = cmdF3.Default
        .Cancel = cmdF3.Cancel
        .Visible = True
    End With
    cmdF3.Visible = False

    With nlcmdF5
        .Init
        .Left = nlcmdF2.Left
        .Top = nlcmdF3.Top + nlcmdF3.Height + 90
        .Caption = cmdF5.Caption
        .TabIndex = cmdF5.TabIndex
        .Enabled = cmdF5.Enabled
        .Default = cmdF5.Default
        .Cancel = cmdF5.Cancel
        .Visible = True
    End With
    cmdF5.Visible = False

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Width = nlcmdF2.Left + nlcmdF2.Width + 300
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxEasyMatch
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
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
'            ElseIf (TypeOf c Is CheckBox) Then
'                c.Height = 0
'                c.Width = c.Height
'                If (c.Name = "chkHistorie") Then
'                    If (c.Index > 0) Then
'                        Load lblchkHistorie(c.Index)
'                    End If
'                    With lblchkHistorie(c.Index)
'                        .BackStyle = 0 'duchsichtig
'                        .Caption = c.Caption
'                        .Left = c.Left + 300
'                        .Top = c.Top
'                        .Width = TextWidth(.Caption) + 90
'                        .TabIndex = c.TabIndex
'                        .Visible = True
'                    End With
'                End If
            End If
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
    nlcmdF3.Visible = False
    nlcmdF5.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Private Sub flxEasyMatch_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEasyMatch_KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    With flxEasyMatch
        If (KeyCode = vbKeyF2) Then
            nlcmdF2.Value = True
        ElseIf (KeyCode = vbKeyF3) Then
            nlcmdF3.Value = True
        ElseIf (KeyCode = vbKeyF5) Then
            nlcmdF5.Value = True
        End If
    End With
Else
    With flxEasyMatch
        If (KeyCode = vbKeyF2) Then
            cmdF2.Value = True
        ElseIf (KeyCode = vbKeyF3) Then
            cmdF3.Value = True
        ElseIf (KeyCode = vbKeyF5) Then
            cmdF5.Value = True
        End If
    End With
End If

Call DefErrPop
End Sub

Sub EditZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim h2$

erg% = EditSatz%(0)
If (erg%) Then erg% = EditSatz%(1)

Call DefErrPop
End Sub

Function EditSatz%(EditCol%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSatz%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim EditRow%, MaxLen%
Dim h2$

If (EditCol% = 0) Then
    EditModus% = 1
    MaxLen% = 40
Else
    EditModus% = 0
    MaxLen% = 10
    If (EasyMatchModus% = 1) Then
        MaxLen% = 8
    End If
End If

EditRow% = flxEasyMatch.row
'EditCol% = flxEasyMatch.col
h2$ = flxEasyMatch.TextMatrix(EditRow%, EditCol%)

Load frmEdit

With frmEdit
    .Left = flxEasyMatch.Left + flxEasyMatch.ColPos(EditCol%)
    If (para.Newline = 0) Then
        .Left = .Left + 45
    End If
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = flxEasyMatch.Top + (EditRow% - flxEasyMatch.TopRow + 1) * flxEasyMatch.RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
    .Width = flxEasyMatch.ColWidth(EditCol%)
    .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit.txtEdit
    .Width = frmEdit.ScaleWidth
    .Left = 0
    .Top = 0
    .text = h2$
    .MaxLength = MaxLen%
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit.Show 1
           
If (EditErg%) Then
    If (Trim$(EditTxt$) = "") Then
        EditErg% = False
    Else
        flxEasyMatch.TextMatrix(EditRow%, EditCol%) = EditTxt$
    End If
End If

EditSatz% = EditErg%

Call DefErrPop
End Function

Sub SpeicherEasyMatch()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherEasyMatch")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, SONDER_HANDLE%, VERBAND%, iPreis%
Dim Name$, Wert$, h$

If (EasyMatchModus% = 0) Then
    On Error Resume Next
    Kill "\user\hmnummer.dat"
    On Error GoTo DefErr
    
    SONDER_HANDLE% = FileOpen("\user\hmnummer.dat", "RW", "B")
    If (SONDER_HANDLE% > 0) Then
        With flxEasyMatch
            For i% = 1 To (.Rows - 1)
                Name$ = Trim(.TextMatrix(i%, 0))
                Call CharToOem(Name$, Name$)
                Wert$ = Trim(.TextMatrix(i%, 1))
                If (Name$ <> "") And (Wert$ <> "") Then
                    h$ = Left$(Wert$ + Space$(10), 10) + Left$(Name$ + Space$(40), 40) + vbCrLf
                    Put #SONDER_HANDLE%, , h$
                End If
            Next i%
        End With
        Close #SONDER_HANDLE%
    End If
ElseIf (EasyMatchModus% = 1) Then
    On Error Resume Next
    Kill "\user\sondrpzn.dat"
    On Error GoTo DefErr
    
    SONDER_HANDLE% = FileOpen("\user\sondrpzn.dat", "RW", "B")
    If (SONDER_HANDLE% > 0) Then
        With flxEasyMatch
            For i% = 1 To (.Rows - 1)
                Name$ = Trim(.TextMatrix(i%, 0))
                Call CharToOem(Name$, Name$)
                Wert$ = Trim(.TextMatrix(i%, 1))
                If (Len(Wert) = 7) Then
                    Wert = "0" + Wert
                End If
                If (Name$ <> "") And (Wert$ <> "") Then
                    h$ = Left$(Wert$ + Space$(8), 8) + Left$(Name$ + Space$(40), 40) + vbCrLf
                    Put #SONDER_HANDLE%, , h$
                End If
            Next i%
        End With
        Close #SONDER_HANDLE%
    End If
Else
    VERBAND% = FileOpen("verbandm.dat", "RW", "B")
    If (VERBAND% > 0) Then
        With flxEasyMatch
            j% = 1
            Seek VERBAND%, 2 * 16 + 1
            For i% = 2 To (.Rows - 1)
                Name$ = Trim(.TextMatrix(i%, 0))
                Wert$ = Trim(.TextMatrix(i%, 1))
                If (Name$ <> "") And (Wert$ <> "") Then
                    Name$ = Left$(Name$ + Space$(12), 12)
                    iPreis% = Val(Wert$)
                    Wert$ = Right$(Space$(4) + Format(iPreis%, "0"), 4)
                    h$ = Name$ + Wert$
                    Put #VERBAND%, , h$
                    j% = j% + 1
                End If
            Next i%
            For i% = j% To 9
                h$ = Space$(16)
                Put #VERBAND%, , h$
            Next i%
        End With
        Close #VERBAND%
    End If
End If

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

Private Sub nlcmdf2_Click()
Call cmdF2_Click
End Sub

Private Sub nlcmdf3_Click()
Call cmdF3_Click
End Sub

Private Sub nlcmdf5_Click()
Call cmdF5_Click
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

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub






