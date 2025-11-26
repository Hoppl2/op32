VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmKKMatch 
   AutoRedraw      =   -1  'True
   Caption         =   "Krankenkassen"
   ClientHeight    =   5025
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   4905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   4905
   Begin nlCommandButton.nlCommand nlcmdKKNeu 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4200
      Picture         =   "frmKKMatch.frx":0000
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
      Left            =   3960
      Picture         =   "frmKKMatch.frx":00B9
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
      Index           =   0
      Left            =   3720
      Picture         =   "frmKKMatch.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdKKNeu 
      Caption         =   "neue Kasse (F2)"
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2760
      TabIndex        =   5
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtKKName 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flxKassen 
      Height          =   2700
      Left            =   360
      TabIndex        =   2
      Top             =   720
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
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblKKName 
      Caption         =   "&Name: "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmKKMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "KKMATCH.FRM"


Sub flxKassenBefüllen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassenBefüllen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
Dim h$, h2$, Sort$, s$, kkey$
Dim ActRecNo&, l&

flxKassen.Rows = 1
With KasseRec
    .index = "Name"
    .MoveFirst
    Do While Not .EOF
        'h$ = Format(.Bookmark, "0")
        h$ = ""
        h$ = h$ + vbTab + KasseRec!Name + vbTab + KasseRec!Nummer + vbTab + KasseRec!kurz
        flxKassen.AddItem h$
        .MoveNext
    Loop
End With

With flxKassen
    .row = 1
    .col = 3
    .RowSel = .Rows - 1
    .ColSel = 1
    .Sort = 5
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
    .SelectionMode = flexSelectionByRow
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

FormErg% = False
Unload Me

Call DefErrPop
End Sub

Private Sub cmdKKNeu_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdKKNeu_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
KKSatz = ""
frmKKassen.Show vbModal
Call flxKassenBefüllen          'Neuanzeige
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

KKSatz = flxKassen.TextMatrix(flxKassen.row, 2)
FormErg% = True
Unload Me

Call DefErrPop
End Sub

Private Sub flxKassen_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_DblClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

Private Sub flxKassen_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With flxKassen
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
    
    txtKKName.text = .TextMatrix(.row, 1)
End With

Call DefErrPop

End Sub


Private Sub flxKassen_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_KeyPress")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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
    With flxKassen
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
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

Private Sub flxKassen_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_RowColChange")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If flxKassen.Visible Then
    If (ActiveControl.Name <> txtKKName.Name) Then
        txtKKName.text = flxKassen.TextMatrix(flxKassen.row, 1)
    End If
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

If (para.Newline) Then
    If KeyCode = vbKeyF2 Then
        Call nlcmdkkneu_Click
    End If
Else
    If KeyCode = vbKeyF2 Then
        Call cmdKKNeu_Click
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control

iEditModus = 1

Me.KeyPreview = True
Call wpara.InitFont(Me)

lblKKName.Top = wpara.TitelY
lblKKName.Left = wpara.LinksX

txtKKName.Left = lblKKName.Left + lblKKName.Width + 150
txtKKName.Top = lblKKName.Top

With flxKassen
    .Rows = 2
    .FixedRows = 1
    .FormatString = "|<Name|<Nummer|<Kurz||"
    .Rows = 1
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth(String(40, "A"))
    .ColWidth(2) = TextWidth(String(12, "9"))
    .ColWidth(3) = TextWidth(String(15, "A"))
    .ColWidth(4) = 0
    .ColWidth(5) = wpara.FrmScrollHeight
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * 13 + 90
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    
    .Top = lblKKName.Top + lblKKName.Height + 150
    txtKKName.Width = .ColWidth(1)
    lblKKName.Visible = True
    txtKKName.Visible = True
    
    Call flxKassenBefüllen
    
End With

Font.Bold = False   ' True

cmdOk.Top = flxKassen.Top + flxKassen.Height + 150
cmdEsc.Top = cmdOk.Top
cmdKKNeu.Top = cmdOk.Top

Me.Width = flxKassen.Left + flxKassen.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height

Me.Font.Name = cmdKKNeu.FontName
Me.Font.Size = cmdKKNeu.FontSize

cmdKKNeu.Width = Me.TextWidth("neue Kasse (F2)") + 300
cmdKKNeu.Height = cmdOk.Height
cmdKKNeu.Left = flxKassen.Left

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxKassen
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
        .Top = flxKassen.Top + flxKassen.Height + iAdd + 600
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

    With nlcmdKKNeu
        .Init
        .AutoSize = True
        .Left = cmdKKNeu.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdKKNeu.Caption
        .TabIndex = cmdKKNeu.TabIndex
        .Enabled = cmdKKNeu.Enabled
        .Default = cmdKKNeu.Default
        .Cancel = cmdKKNeu.Cancel
        .Visible = True
    End With
    cmdKKNeu.Visible = False

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxKassen
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
    nlcmdKKNeu.Visible = False
End If
'''''''''
Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Me.Font.Name = wpara.FontName(0)
Me.Font.Size = wpara.FontSize(0)

flxKassen.col = 2
flxKassen.col = 1

Call DefErrPop

End Sub


Private Sub txtKKName_Change()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtKKName_Change")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, l%
Dim h$

h$ = UCase$(Trim(txtKKName.text))
l% = Len(h$)

With flxKassen
    For i% = 1 To (.Rows - 1)
        If (Left$(.TextMatrix(i%, 1), l%) = h$) Then
            .TopRow = i%
            .row = i%
            Exit For
        End If
    Next i%
End With

Call DefErrPop

End Sub


Private Sub txtKKName_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtKKName_GotFocus")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

With txtKKName
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop

End Sub


Private Sub txtKKName_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtKKName_KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If KeyCode = vbKeyDown Then
    If flxKassen.row < flxKassen.Rows - 1 Then
        flxKassen.row = flxKassen.row + 1
        txtKKName.text = flxKassen.TextMatrix(flxKassen.row, 1)
        If flxKassen.row + 1 < flxKassen.Rows Then
          While Not (flxKassen.RowIsVisible(flxKassen.row + 1))
              flxKassen.TopRow = flxKassen.TopRow + 1
          Wend
        ElseIf flxKassen.TopRow < flxKassen.Rows - 1 Then
            flxKassen.TopRow = flxKassen.TopRow + 1
        End If
        flxKassen.SetFocus
    End If
    KeyCode = 0
'ElseIf KeyCode = vbKeyUp Then
'    If flxKassen.row > 1 Then
'        flxKassen.row = flxKassen.row - 1
'        txtKKName.text = flxKassen.TextMatrix(flxKassen.row, 1)
 '       If flxKassen.row < flxKassen.TopRow Then
'            flxKassen.TopRow = flxKassen.row
'        End If
'    End If
'    KeyCode = 0
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

Private Sub nlcmdkkneu_Click()
Call cmdKKNeu_Click
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






