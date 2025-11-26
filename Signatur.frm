VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmSignatur 
   Caption         =   "Benutzer-Signatur"
   ClientHeight    =   3045
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4395
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4395
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3360
      Picture         =   "Signatur.frx":0000
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
      Left            =   3600
      Picture         =   "Signatur.frx":00A9
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
      Left            =   3840
      Picture         =   "Signatur.frx":015D
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
      Height          =   450
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1200
   End
   Begin VB.TextBox txtSignatur 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   390
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblSignatur 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie Ihr &Passwort ein:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmSignatur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "SIGNATUR.FRM"

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
Dim iBenutzerInd%, okAktiv%

'If (iNewLine) Then
'    okAktiv = (ActiveControl.Name = nlcmdOk.Name)
'Else
'    okAktiv = (ActiveControl.Name = cmdOk.Name)
'End If
'If (okAktiv) Then
    iBenutzerInd% = CheckPasswort%
    If (iBenutzerInd% > 0) Then
        SignaturEingabeErg% = iBenutzerInd%
        Unload Me
    Else
        txtSignatur.SetFocus
    End If
'Else
'    clsSI.MySendKeys "{TAB}", True
'End If

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
Dim iAdd%, iAdd2%, x%, y%, wi%, ydiff%

SignaturEingabeErg% = 0

Call wPara1.InitFont(Me)

With lblSignatur
    .Left = wPara1.LinksX
    .Top = 2 * wPara1.TitelY
End With
With txtSignatur
    ydiff% = (.Height - lblSignatur.Height) / Screen.TwipsPerPixelY
    ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
    .Top = lblSignatur.Top - ydiff%
    .Left = lblSignatur.Left + lblSignatur.Width + 300
    .Width = TextWidth(String(15, "X"))
End With

Me.Width = txtSignatur.Left + txtSignatur.Width + 3 * wPara1.LinksX

With cmdOk
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    .Top = lblSignatur.Top + lblSignatur.Height + 600
End With

With cmdEsc
    .Width = cmdOk.Width
    .Height = cmdOk.Height
    .Left = cmdOk.Left + cmdEsc.Width + 300
    .Top = cmdOk.Top
End With

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.FrmCaptionHeight + 2 * wPara1.TitelY


''''''
If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    lblSignatur.Top = lblSignatur.Top + iAdd2
    txtSignatur.Top = txtSignatur.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdOk.Top
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = txtSignatur.Top + txtSignatur.Height + iAdd + 600
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
        .Left = nlcmdOk.Left + .Width + 300
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxPosPartner.Left - iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top - iAdd) / Screen.TwipsPerPixelY, (flxPosPartner.Left + flxPosPartner.Width + iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top + flxPosPartner.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    With txtSignatur
'        .Appearance = 0
        .BackColor = vbWhite
        Call wPara1.ControlBorderless(txtSignatur, 1, 1)
    End With
    With Me
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtSignatur.Left - 90) / Screen.TwipsPerPixelX, (txtSignatur.Top - 45) / Screen.TwipsPerPixelY, (txtSignatur.Left + txtSignatur.Width + 90) / Screen.TwipsPerPixelX, (txtSignatur.Top + txtSignatur.Height + 45) / Screen.TwipsPerPixelY, 10, 10
    End With

'    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
'    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If


''''''''

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

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
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
    With txtSignatur
'        .Appearance = 0
        .BackColor = vbWhite
        Call wPara1.ControlBorderless(txtSignatur, 1, 1)
    End With
    With Me
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtSignatur.Left - 90) / Screen.TwipsPerPixelX, (txtSignatur.Top - 45) / Screen.TwipsPerPixelY, (txtSignatur.Left + txtSignatur.Width + 90) / Screen.TwipsPerPixelX, (txtSignatur.Top + txtSignatur.Height + 45) / Screen.TwipsPerPixelY, 10, 10
    End With

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub txtSignatur_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtSignatur_GotFocus")
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

With txtSignatur
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Function CheckPasswort%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckPasswort%")
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
Dim i%, ret%
Dim pass$, GeneralPW$

ret% = 0

pass$ = UCase(Trim(txtSignatur.text))
If (pass$ <> "") Then
    GeneralPW = "@" + Format(Now, "hhnnddmm")        '1.0.77
    If (pass = GeneralPW) Then
'        Para1.Passwort(1) = "OPTIPHARM"
        ret = 1
    Else
        For i% = 1 To 80
            If (pass$ = Para1.Passwort(i%)) Then
                ret% = i%
                Exit For
            End If
        Next i%
    End If
End If

CheckPasswort% = ret%

Call clsError.DefErrPop
End Function

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

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub




