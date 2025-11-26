VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmAvDebug 
   AutoRedraw      =   -1  'True
   Caption         =   "A+V  Debug"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8475
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   7320
      Picture         =   "avdebug.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   7560
      Picture         =   "avdebug.frx":00A9
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
      Index           =   2
      Left            =   7800
      Picture         =   "avdebug.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
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
      Left            =   2640
      TabIndex        =   3
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAvDebug 
      Height          =   2700
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   840
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
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxAvDebug 
      Height          =   2640
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
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
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxAvDebug 
      Height          =   2640
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
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
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmAvDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "AVDEBUG.FRM"

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

Private Sub flxAvDebug_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvDebug_GotFocus")
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

With flxAvDebug(index)
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxAvDebug_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvDebug_LostFocus")
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

With flxAvDebug(index)
    .HighLight = flexHighlightNever
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control

Call wpara.InitFont(Me)
'Me.Font.Size = wpara.FontSize(0)

iEditModus = 1

With flxAvDebug(0)
    .Rows = 1
    .Cols = 2
    .ColWidth(0) = TextWidth(String(12, "X"))
    .ColWidth(1) = TextWidth(String(31, "9"))
    .ColAlignment(1) = flexAlignLeftCenter
    .Width = .ColWidth(0) + .ColWidth(1) + 90
    .Height = .RowHeight(0) * 18 + 90
    .Rows = 0
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
End With

With flxAvDebug(1)
    .Rows = 2
    .FixedRows = 1
    .FormatString = "|>Key/Bed|^WaZ|^Rnd|^Teil|^Neg|^Erm|^KV|^GPfl|^Pau|^Zuz|^Abr|^HmAbrKz|^Verw.Bed"
    .Rows = 1
    .ColWidth(0) = TextWidth(String(9, "X"))
    .ColWidth(1) = TextWidth(String(37, "9"))
    For i% = 2 To (.Cols - 2)
        .ColWidth(i%) = TextWidth("XXXX")
    Next i%
    .ColWidth(.Cols - 2) = TextWidth(String(8, "X"))
    .ColWidth(.Cols - 1) = TextWidth(String(8, "X"))
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * 13 + 90
    
    .Top = wpara.TitelY
    .Left = flxAvDebug(0).Left + flxAvDebug(0).Width + 150
End With

With flxAvDebug(2)
    .Rows = 2
    .FixedRows = 1
    .FormatString = "|>Berech/Wert|^Mw|^Rab|>Berech/Wert|^Mw|^Rab|>Berech/Wert|^Mw|^Rab|>Berech/Wert|^Mw|^Rab"
    .Rows = 1
    
    .ColWidth(0) = TextWidth(String(17, "X"))
    For i% = 0 To 2
        ind% = i% * 3 + 1
        If (i% = 0) Then
            .ColWidth(ind%) = TextWidth(String(35, "9"))
        Else
            .ColWidth(ind%) = TextWidth(String(24, "9"))
        End If
        .ColWidth(ind% + 1) = TextWidth(String(4, "X"))
        .ColWidth(ind% + 2) = TextWidth(String(4, "X"))
    Next i%
    i% = 3
    ind% = i% * 3 + 1
    .ColWidth(ind%) = 0
    .ColWidth(ind% + 1) = 0
    .ColWidth(ind% + 2) = 0
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * 8 + 90
    
    .Top = flxAvDebug(0).Top + flxAvDebug(0).Height + 90
    .Left = flxAvDebug(0).Left
End With

Font.Bold = False   ' True

cmdEsc.Top = flxAvDebug(2).Top + flxAvDebug(2).Height + 150

Me.Width = flxAvDebug(1).Left + flxAvDebug(1).Width + 2 * wpara.LinksX

cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEsc.Left = (Me.Width - cmdEsc.Width) / 2

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    For j% = 0 To 2
        With flxAvDebug(j%)
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
    Next j
    
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
        .Top = flxAvDebug(2).Top + flxAvDebug(2).Height + iAdd + 600
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
    With flxAvDebug(0)
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (flxAvDebug(1).Left + flxAvDebug(1).Width + iAdd) / Screen.TwipsPerPixelX, (flxAvDebug(2).Top + flxAvDebug(2).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
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






