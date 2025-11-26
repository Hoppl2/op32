VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmWechsel 
   AutoRedraw      =   -1  'True
   Caption         =   "Wechsel"
   ClientHeight    =   2985
   ClientLeft      =   1785
   ClientTop       =   3195
   ClientWidth     =   3480
   Icon            =   "Wechsel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3480
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   2040
      Picture         =   "Wechsel.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "Wechsel.frx":0203
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   1560
      Picture         =   "Wechsel.frx":02B7
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox lstWechsel 
      Height          =   1635
      Left            =   0
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmWechsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WECHSEL.FRM"

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
Dim i%

With lstWechsel
    WechselErg% = -1
    WechselStart% = 0
    For i% = 0 To .ListCount - 1
        If (.Selected(i%)) Then
            WechselStart% = i%
            Exit For
        End If
    Next i%
End With

Unload Me

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
Dim i%

With lstWechsel
    WechselErg% = .ListIndex
    WechselStart% = 0
    For i% = 0 To .ListCount - 1
        If (.Selected(i%)) Then
            WechselStart% = i%
            Exit For
        End If
    Next i%
End With

Unload Me

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
Dim i%, MaxWi%, AnzRows%
Dim iAdd%, iAdd2%, x%, y%, wi%, ydiff%
Dim h$
                
Call wPara1.InitFont(Me)

With lstWechsel
    .Clear
    MaxWi% = 0
    AnzRows% = UBound(WechselZeilen$) + 1
    For i% = 0 To AnzRows% - 1
        h$ = WechselZeilen$(i%)
        .AddItem h$
        wi% = TextWidth(h$)
        If (wi% > MaxWi%) Then
            MaxWi% = wi%
        End If
    Next i%
    .Selected(WechselStart%) = True
    .ListIndex = WechselStart%
End With

Width = MaxWi% + 900
Height = AnzRows% * (TextHeight("Äg") * 1.3) + 90 + 2 * wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
'Left = frmMatchcode.Left + (frmMatchcode.Width - Width) / 2
'Top = frmMatchcode.Top + (frmMatchcode.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
    
With lstWechsel
    .Height = ScaleHeight
    .Width = ScaleWidth
    .Left = 0
    .Top = 0
    Height = .Height + 2 * wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
End With

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With lstWechsel
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdOk.Top
'
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    lstWechsel.Top = lstWechsel.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdOk.Top
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = 0
        .Top = -900
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .default = cmdOk.default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
        .Left = 0
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

'    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450
    
    Call wPara1.NewLineWindow(Me, Height)
'    BackColor = vbGreen ' RGB(180, 180, 180)
'
'    Me.Line (15, 15)-(ScaleWidth - 30, 180), RGB(247, 247, 247), BF
'
'    ForeColor = RGB(177, 177, 177)
'    FillStyle = vbSolid
'    FillColor = RGB(247, 247, 247)
'    RoundRect hdc, 0, 0, ScaleWidth / Screen.TwipsPerPixelX, ScaleHeight / Screen.TwipsPerPixelY, 20, 20
'
'    Call wPara1.FillGradient(Me, 1, 180 / Screen.TwipsPerPixelY, ScaleWidth / Screen.TwipsPerPixelX - 2, 450 / Screen.TwipsPerPixelY, RGB(230, 230, 230), RGB(191, 191, 191))
'
''    Y = nlcmdOk.Top + 210
''    Me.Line (15, Y)-(ScaleWidth - 30, ScaleHeight - 15), RGB(177, 177, 177), BF
'    y = Height
'    Call wPara1.FillGradient(Me, 1, 450 / Screen.TwipsPerPixelY, ScaleWidth / Screen.TwipsPerPixelX - 2, y / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
'
''    Y = nlcmdOk.Top - 300
''    Me.ForeColor = RGB(177, 177, 177)
''    Me.Line (15, Y)-(ScaleWidth - 15, Y)
'
'    wi = 16 * Screen.TwipsPerPixelX
'    x = ScaleWidth - 210 - wi
'    y = (450 - wi) / 2
'    If (Me.MinButton) Then
'        With picControlBox(0)
'            .Left = x - 2 * (wi + 45)
'            .Top = y
'            .Visible = True
''            Me.PaintPicture ProjektForm.imgToolbar(2).ListImages(20).Picture, x - 2 * (wi + 45), Y, wi, wi
'        End With
'    End If
'    If (Me.MaxButton) Then
'        With picControlBox(1)
'            .Left = x - (wi + 45)
'            .Top = y
'            .Visible = True
''            Me.PaintPicture ProjektForm.imgToolbar(2).ListImages(20).Picture, x - (wi + 45), Y, wi, wi
'        End With
'    End If
'    If (Me.ControlBox) Then
'        With picControlBox(2)
'            .Left = x
'            .Top = y
'            .Visible = True
''            Me.PaintPicture ProjektForm.imgToolbar(2).ListImages(20).Picture, x, Y, wi, wi
'        End With
'    End If
'
'    Height = Height + wPara1.FrmCaptionHeight + 1
'    Call wPara1.ControlBorderless(Me, -3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 4)
    
    RoundRect hdc, (lstWechsel.Left - iAdd) / Screen.TwipsPerPixelX, (lstWechsel.Top - iAdd) / Screen.TwipsPerPixelY, (lstWechsel.Left + lstWechsel.Width + iAdd) / Screen.TwipsPerPixelX, (lstWechsel.Top + lstWechsel.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

'    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
'    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
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
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    Call wPara1.NewLineWindow(Me, Height, False)
    RoundRect hdc, (lstWechsel.Left - iAdd) / Screen.TwipsPerPixelX, (lstWechsel.Top - iAdd) / Screen.TwipsPerPixelY, (lstWechsel.Left + lstWechsel.Width + iAdd) / Screen.TwipsPerPixelX, (lstWechsel.Top + lstWechsel.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub lstWechsel_ItemCheck(Item As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lstWechsel_ItemCheck")
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
Static Aktiv%

If (Aktiv%) Then Call clsError.DefErrPop: Exit Sub

Aktiv% = True

With lstWechsel
    For i% = 0 To .ListCount - 1
        If (i% <> Item) Then
            .Selected(i%) = False
        End If
    Next i%
End With

Aktiv% = False

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
    CurrentX = wPara1.NlFlexBackY
    CurrentY = (wPara1.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
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
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
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





