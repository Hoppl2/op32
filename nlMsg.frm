VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmnlMsg 
   AutoRedraw      =   -1  'True
   Caption         =   "Benutzer-Signatur"
   ClientHeight    =   3045
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4395
   Begin VB.PictureBox picTemp 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3360
      Picture         =   "nlMsg.frx":0000
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
      Picture         =   "nlMsg.frx":00A9
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
      Picture         =   "nlMsg.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtMsg 
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
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   390
      Visible         =   0   'False
      Width           =   855
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin ComctlLib.ImageList imgMsg 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   75
      ImageHeight     =   75
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlMsg.frx":0216
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlMsg.frx":1CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "nlMsg.frx":3742
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmnlMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HochfahrenAktiv%

Private Const DefErrModul = "NLMSG.FRM"

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

If (nlmsgVisible(2)) Then
    nlmsgRet2 = Trim(txtMsg)
Else
    nlmsgRet = vbYes
End If

Unload Me

Call clsError.DefErrPop
End Sub

Private Sub txtMsg_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtMsg_GotFocus")
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

With txtMsg
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtMsg_KeyPress")
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

If (nlmsgEditModus <> 1) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((nlmsgEditModus% <> 2) And (nlmsgEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
        Beep
        KeyAscii = 0
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Unload")
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

nlmsgAktiv = 0

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
Dim iAdd%, iAdd2%, x%, y%, wi%
Dim i%, j%, k%, m%, ff%, PixelOk%
Dim lColor&
Dim h$, h2$
Dim c As Control

'SignaturEingabeErg% = 0

HochfahrenAktiv% = True

nlmsgAktiv = True

Call wPara1.InitFont(Me)

Me.Caption = nlmsgTitle

nlmsgRet = vbNo
nlmsgRet2 = ""

With lblMsg
    .Left = 300
    If (nlmsgPicto >= 0) Then
        .Left = .Left + 1125 + 300
    End If
    .Top = 2 * wPara1.TitelY
    
'    .Width = 3000
    .Caption = nlmsgPrompt + vbCrLf
    .FontBold = True
    .Width = TextWidth(.Caption) + 300
    .Height = TextHeight(.Caption) + 300
'    .Visible = False
End With

Me.Width = lblMsg.Left + lblMsg.Width + 3 * wPara1.LinksX

txtMsg.Tag = "0"
If (nlmsgVisible(2)) Then
    With txtMsg
        .Left = lblMsg.Left + lblMsg.Width + 300
        .Top = lblMsg.Top
        .text = nlmsgCaption(2)
        .Width = TextWidth(String(15, "X")) '+ 300
        .Tag = ""
        If (nlmsgEditModus = 999) Then
            .PasswordChar = "*"
            nlmsgEditModus = 1
            .text = ""
        End If
        .Visible = True
    End With
    Me.Width = txtMsg.Left + txtMsg.Width + 3 * wPara1.LinksX
End If

iAdd = wPara1.NlFlexBackY
iAdd2 = wPara1.NlCaptionY

lblMsg.Top = lblMsg.Top + iAdd2
txtMsg.Top = txtMsg.Top + iAdd2
    
With nlcmdOk
    .Init
    wi = 2 * (.Width + 300) + 300
    If (wi > Me.Width) Then
        Me.Width = wi
    End If
    If (nlmsgVisible(1)) Then
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
'        .Left = (Me.ScaleWidth - 2 * (.Width + 300))
    Else
        .Left = (Me.ScaleWidth - (.Width)) / 2
'        .Left = (Me.ScaleWidth - (.Width + 300))
    End If
    
    y = 1125
    If (lblMsg.Height > 1125) Then
        y = lblMsg.Height
    End If
    .Top = lblMsg.Top + y + iAdd + 600
    .Caption = nlmsgCaption(0)
    .default = nlmsgDefault(0)
    .Cancel = nlmsgCancel(0)
    .Visible = True
End With

With nlcmdEsc
    .Init
    .Left = nlcmdOk.Left + .Width + 300
    .Top = nlcmdOk.Top
    .Caption = nlmsgCaption(1)
    .Enabled = True
    .default = nlmsgDefault(1)
    .Cancel = nlmsgCancel(1)
    .Visible = nlmsgVisible(1)
End With

Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450
    
'''''
Call wPara1.NewLineWindow(Me, nlcmdEsc.Top)

If (nlmsgPicto >= 0) Then
    With picTemp
        .BorderStyle = 0
        .Left = ScaleWidth - 3000
        .Top = 1600
        
        .BackColor = vbWhite

        .Width = 1125
        .Height = 1125  'IconHeight%
        .AutoRedraw = True
        .Cls
        .PaintPicture imgMsg.ListImages(nlmsgPicto + 1).Picture, 0, 0, 1125, 1125

'                ff% = clsDat.FileOpen("\user\pixel.txt", "O")
        m = 74
        y = lblMsg.Top / Screen.TwipsPerPixelY
        For j = 2 To m
            For k = 2 To m
                lColor = GetPixel(.hdc, j, k)
                If (lColor <> .BackColor) Then
                    SetPixel Me.hdc, 20 + j, y + k, lColor
                    
'                    PixelOk = 0
'                    h = Right("000000" + Hex(lColor), 6)
'                    For i = 0 To 2
'                        If (Val("&H" + Left(h, 2)) < 250) Then
'                            PixelOk = True
'                            Exit For
'                        End If
'                        h = Mid$(h, 3)
'                    Next i
'                    If (PixelOk) Then
'                        SetPixel Me.hdc, 20 + j + 100, y + k, lColor
'                    End If
                End If
            Next k
        Next j
        
'        Close ff%
    End With
End If
    
    On Error Resume Next
    For Each c In Controls
        If (c.Tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (c.BorderStyle > 0) Then
                    If (TypeOf c Is ComboBox) Then
                        Call wPara1.ControlBorderless(c)
                    ElseIf (c.Appearance = 1) Then
                        Call wPara1.ControlBorderless(c, 2, 2)
                    Else
                        Call wPara1.ControlBorderless(c, 1, 1)
                    End If
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
            ElseIf (TypeOf c Is CheckBox) Then
                c.Height = 0
                c.Width = c.Height
'                If (c.Name = "chkInStamm") Then
'                    If (c.index > 0) Then
'                        Load lblchkInStamm(c.index)
'                    End If
'                    With lblchkInStamm(c.index)
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
''''

Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

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
    
If (y <= 450) Then
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

If (HochfahrenAktiv%) And (Me.Visible) Then
    HochfahrenAktiv = 0
    If (txtMsg.Visible) Then
        txtMsg.SetFocus
'    ElseIf (nlcmdOk.Visible) Then
'        nlcmdOk.SetFocus
    ElseIf (nlcmdEsc.Visible) And (nlcmdEsc.default) Then
        nlcmdEsc.SetFocus
    Else
        nlcmdOk.SetFocus
    End If
    If Not (nlmsgVisible(2)) Then
        nlcmdOk.default = False
        nlcmdEsc.default = False
    End If
End If

If (iNewLine) And (Me.Visible) Then
    CurrentX = 210
    CurrentY = (450 - TextHeight(Caption)) / 2
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyDown")
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
Dim ind%
Dim c As Control

If (Shift And vbAltMask) Then
    If (KeyCode <> 18) Then
        On Error Resume Next
        For Each c In Controls
            If (TypeOf c Is nlCommand) Then
                If (c.Accelerator <> "") And (c.Enabled) Then
                    If (Asc(c.Accelerator) = KeyCode) Then
                        c.SetFocus
                        c.Value = 1
'                        Exit For
                        Call clsError.DefErrPop: Exit Sub
                    End If
                End If
            End If
        Next
        On Error GoTo DefErr
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim ch$

ch = Chr(KeyAscii)
If (iNewLine) Then
    If (KeyAscii = 13) Then
        If (nlmsgVisible(2)) Then
            Call nlcmdOk_Click
        Else
            ActiveControl.Value = True
        End If
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    ElseIf Not (nlmsgVisible(2)) Then
        If (ch = "j") Or (ch = "J") Then
            Call nlcmdOk_Click
        ElseIf (nlmsgVisible(1)) And ((ch = "n") Or (ch = "N")) Then
            Call nlcmdEsc_Click
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




