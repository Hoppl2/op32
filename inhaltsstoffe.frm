VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmInhaltsStoffe 
   AutoRedraw      =   -1  'True
   Caption         =   "Inhaltsstoffe zu"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5895
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   5280
      Picture         =   "inhaltsstoffe.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   5040
      Picture         =   "inhaltsstoffe.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   4800
      Picture         =   "inhaltsstoffe.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   1
      Top             =   3720
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxInhaltsStoffe 
      Height          =   2280
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblInhaltsStoffe 
      Caption         =   "Bitte prüfen Sie die folgenden Verordnungen und deren Regelungen:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmInhaltsStoffe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "INHALTSSTOFFE.FRM"

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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%
Dim DRUCKHANDLE%
Dim iAdd%, iAdd2%
Dim h$, h2$
Dim c As Control

iEditModus = 1

Call wpara.InitFont(Me)

With lblInhaltsStoffe
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Visible = (InhaltsstoffeModus = 1)
End With

With flxInhaltsStoffe
    .Cols = 2
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    If (InhaltsstoffeModus = 0) Then
        .FormatString = "<Inhaltsstoff|>Menge"
        .Rows = 1
        
        Call ActProgram.FlxInhaltsStoffeBefuellen
        
        .Top = wpara.TitelY
        .Left = wpara.LinksX
        .Height = .RowHeight(0) * .Rows + 90
        
        .ColAlignment(0) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth(String(35, "X"))
        .ColWidth(1) = TextWidth(String(18, "9"))
        .Width = .ColWidth(0) + .ColWidth(1) + 90
    Else
        .FormatString = "<Vdb-Artikel|<Regelung"
        .Rows = 2
        
        .Top = lblInhaltsStoffe.Top + lblInhaltsStoffe.Height + 210
        .Left = wpara.LinksX
        .Height = .RowHeight(0) * 16 + 90
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .SelectionMode = flexSelectionByRow
        .BackColor = vbButtonFace
        
        .ColWidth(0) = TextWidth(String(35, "X"))
        .ColWidth(1) = TextWidth(String(35, "9"))
        .Width = .ColWidth(0) + .ColWidth(1) + 90
    End If
        
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxInhaltsStoffe.Left + flxInhaltsStoffe.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = flxInhaltsStoffe.Top + flxInhaltsStoffe.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxInhaltsStoffe
        .ScrollBars = flexScrollBarNone
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
        .Top = flxInhaltsStoffe.Top + flxInhaltsStoffe.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    nlcmdEsc.Left = (Me.Width - nlcmdEsc.Width) / 2

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxInhaltsStoffe
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
    nlcmdEsc.Visible = False
End If
'''''''''

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

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

'Private Sub nlcmdOk_Click()
'Call cmdOk_Click
'End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
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







