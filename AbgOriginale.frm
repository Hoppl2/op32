VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmAbgOriginale 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Parameter für Abgegebene Originale"
   ClientHeight    =   4230
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7725
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   7200
      Picture         =   "AbgOriginale.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   6960
      Picture         =   "AbgOriginale.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   6720
      Picture         =   "AbgOriginale.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtImpAlternativ 
      Height          =   375
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "99999999"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtImpAlternativ 
      Height          =   375
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "99999999"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtImpAlternativ 
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   6
      Text            =   "99999999"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   4800
      TabIndex        =   9
      Top             =   2520
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblImpAlternativ2 
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblImpAlternativ 
      Caption         =   "Abgegebene Originale auf Kassenrezept &seit:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblImpAlternativ 
      Caption         =   "&ab:"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblImpAlternativ2 
      Caption         =   "Packungen"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblImpAlternativ 
      Caption         =   "&oder:"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblImpAlternativ2 
      Caption         =   "EUR Avp-Gesamtwert"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbgOriginale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "ABGORIGINALE.FRM"

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
Dim i%
Dim AbDatum$

If (ActiveControl.Name = txtImpAlternativ(0).Name) Then
    MySendKeys "{TAB}", True
Else
    AbDatum$ = Trim(txtImpAlternativ(0).text)
    If (iDate(AbDatum$) = 0) Then
        txtImpAlternativ(0).text = "0101" + Format(Now, "YY")
        txtImpAlternativ(0).SetFocus
    Else
        EditErg% = True
        For i% = 0 To 2
            ImpAlternativPara$(i%) = Trim(txtImpAlternativ(i%).text)
        Next i%
        Unload Me
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
Dim i%, Breite%, MaxWi%, wi%, diff%, FormVersatzY%
Dim iAdd%, iAdd2%
Dim c As Control

iEditModus = 1

EditErg% = 0

Call wpara.InitFont(Me)

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr

txtImpAlternativ(0).text = "0101" + Format(Now, "YY")
txtImpAlternativ(1).text = "1"

txtImpAlternativ(0).Top = 2 * wpara.TitelY
For i% = 1 To 2
    txtImpAlternativ(i%).Top = txtImpAlternativ(i% - 1).Top + txtImpAlternativ(i% - 1).Height + 90
Next i%

diff% = (txtImpAlternativ(0).Height - lblImpAlternativ(0).Height) / 2
lblImpAlternativ(0).Left = wpara.LinksX
lblImpAlternativ(0).Top = txtImpAlternativ(0).Top + diff%
For i% = 1 To 2
    lblImpAlternativ(i%).Left = lblImpAlternativ(0).Left + lblImpAlternativ(0).Width - lblImpAlternativ(i%).Width
    lblImpAlternativ(i%).Top = txtImpAlternativ(i%).Top + diff%
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblImpAlternativ(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtImpAlternativ(0).Left = lblImpAlternativ(0).Left + MaxWi% + 300
For i% = 1 To 2
    txtImpAlternativ(i%).Left = txtImpAlternativ(i% - 1).Left
Next i%

lblImpAlternativ2(0).Left = txtImpAlternativ(0).Left + txtImpAlternativ(0).Width + 150
lblImpAlternativ2(0).Top = txtImpAlternativ(0).Top
For i% = 1 To 2
    lblImpAlternativ2(i%).Left = lblImpAlternativ2(i% - 1).Left
    lblImpAlternativ2(i%).Top = txtImpAlternativ(i%).Top
Next i%

'With fmeImpAlternativ
'    .Left = wpara.LinksX
'    .Top = wpara.TitelY
'
'    .Width = lblImpAlternativ2(2).Left + lblImpAlternativ2(2).Width + 2 * wpara.LinksX
'    .Height = txtImpAlternativ(2).Top + txtImpAlternativ(2).Height + 2 * wpara.TitelY
'End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

'Me.Width = fmeImpAlternativ.Left + fmeImpAlternativ.Width + 2 * wpara.LinksX
Me.Width = lblImpAlternativ2(2).Left + lblImpAlternativ2(2).Width + 2 * wpara.LinksX

With cmdOk
'    .Top = fmeImpAlternativ.Top + fmeImpAlternativ.Height + 150 * wpara.BildFaktor
    .Top = txtImpAlternativ(2).Top + txtImpAlternativ(2).Height + 450
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
End With
With cmdEsc
    .Top = cmdOk.Top
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = cmdOk.Left + cmdEsc.Width + 300
End With


Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
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
        .Top = txtImpAlternativ(2).Top + txtImpAlternativ(2).Height + iAdd + 600
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

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)

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
End If
'''''''''

Me.Left = frmRezSpeicher.Left + (frmRezSpeicher.Width - Me.Width) / 2
If (Me.Left < 0) Then
    Me.Left = 0
End If

Me.Top = frmRezSpeicher.Top + (frmRezSpeicher.Height - Me.Height) / 2
If (Me.Top < 0) Then
    Me.Top = 0
End If

Call DefErrPop
End Sub

Private Sub txtImpAlternativ_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtImpAlternativ_GotFocus")
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

With txtImpAlternativ(index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtImpAlternativ_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtImpAlternativ_KeyPress")
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

If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
    Beep
    KeyAscii = 0
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






