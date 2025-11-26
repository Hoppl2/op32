VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmSubstitutionsRezept 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Parameter für Substitutionsrezept"
   ClientHeight    =   5340
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7725
   Begin VB.CheckBox chkSubstitutionsrezept 
      Caption         =   "mit &kindergesichertem Verschluss"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Tag             =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox txtSubstitutionsrezept 
      Height          =   375
      Index           =   3
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   10
      Text            =   "99999999"
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   7200
      Picture         =   "SubstitutionsRezept.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
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
      Picture         =   "SubstitutionsRezept.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
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
      Picture         =   "SubstitutionsRezept.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtSubstitutionsrezept 
      Height          =   375
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "99999999"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtSubstitutionsrezept 
      Height          =   375
      Index           =   1
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "99999999"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtSubstitutionsrezept 
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   7
      Text            =   "99999999"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   2880
      TabIndex        =   12
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   4440
      TabIndex        =   13
      Top             =   3600
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   4440
      TabIndex        =   17
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblchkSubstitutionsrezept 
      Caption         =   "AAAA"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   20
      Tag             =   "0"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblSubstitutionsrezept2 
      Caption         =   "Einzeldosen"
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblSubstitutionsrezept 
      Caption         =   "&3.Abgabe"
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label lblSubstitutionsrezept2 
      Caption         =   "mg"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSubstitutionsrezept 
      Caption         =   "&Einzeldosis:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblSubstitutionsrezept 
      Caption         =   "&1.Abgabe:"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblSubstitutionsrezept2 
      Caption         =   "Einzeldosen"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSubstitutionsrezept 
      Caption         =   "&2.Abgabe:"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblSubstitutionsrezept2 
      Caption         =   "Einzeldosen"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmSubstitutionsRezept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "SubstitutionsRezept.FRM"

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
Dim i%, iVal%, iOk%
Dim AbDatum$

If (ActiveControl.Name = txtSubstitutionsrezept(0).Name) Then
    MySendKeys "{TAB}", True
Else
    iOk = True
    If (ParenteralRezept = 20) Or (ParenteralRezept = 21) Or (ParenteralRezept = 22) Or (ParenteralRezept = 23) Then
        If (Val(txtSubstitutionsrezept(0).text) > 24) Then
            Call MessageBox("Achtung: Tagesdosis maximal 24 mg!", vbCritical)
            txtSubstitutionsrezept(0).SetFocus
            iOk = 0
        Else
            iVal = 0
            For i = 0 To 2
                iVal = iVal + Val(txtSubstitutionsrezept(i + 1).text)
            Next i
            If (iVal > 7) Then
                If (MessageBox("Achtung: Maximal 7 Einzeldosen!" + vbCrLf + vbCrLf + "Rezeptur dennoch speichern?", vbYesNo Or vbInformation Or vbDefaultButton2) <> vbYes) Then
                    txtSubstitutionsrezept(1).SetFocus
                    iOk = 0
                End If
            End If
        End If
    End If
    If (iOk) Then
        For i = 0 To 2
            If (Val(txtSubstitutionsrezept(i + 1).text) > 30) Then
                Call MessageBox("Achtung: Max. 30 Einzeldosen pro Abgabe!", vbCritical)
                txtSubstitutionsrezept(i + 1).SetFocus
                iOk = 0
            End If
        Next i
    End If
    If (iOk) Then
        SubstitutionsMg = Val(txtSubstitutionsrezept(0).text)
        For i = 0 To 2
            SubstitutionsAbgaben(i) = Val(txtSubstitutionsrezept(i + 1).text)
        Next i
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

'txtSubstitutionsrezept(0).text = "0101" + Format(Now, "YY")
'txtSubstitutionsrezept(1).text = "1"

txtSubstitutionsrezept(0).text = IIf(SubstitutionsMg > 0, CStr(SubstitutionsMg), "")
For i = 0 To 2
    txtSubstitutionsrezept(i + 1).text = IIf(SubstitutionsAbgaben(i) > 0, CStr(SubstitutionsAbgaben(i)), "")
Next i


txtSubstitutionsrezept(0).Top = 2 * wpara.TitelY
For i% = 1 To 3
    txtSubstitutionsrezept(i%).Top = IIf(i = 1, 300, 0) + txtSubstitutionsrezept(i% - 1).Top + txtSubstitutionsrezept(i% - 1).Height + 60
Next i%

diff% = (txtSubstitutionsrezept(0).Height - lblSubstitutionsrezept(0).Height) / 2
lblSubstitutionsrezept(0).Left = 2 * wpara.LinksX
lblSubstitutionsrezept(0).Top = txtSubstitutionsrezept(0).Top + diff%
For i% = 1 To 3
    lblSubstitutionsrezept(i%).Left = lblSubstitutionsrezept(0).Left '+ lblSubstitutionsrezept(0).Width - lblSubstitutionsrezept(i%).Width
    lblSubstitutionsrezept(i%).Top = txtSubstitutionsrezept(i%).Top + diff%
Next i%

MaxWi% = 0
For i% = 0 To 3
    wi% = lblSubstitutionsrezept(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtSubstitutionsrezept(0).Left = lblSubstitutionsrezept(0).Left + MaxWi% + 600
For i% = 1 To 3
    txtSubstitutionsrezept(i%).Left = txtSubstitutionsrezept(i% - 1).Left
Next i%

lblSubstitutionsrezept2(0).Left = txtSubstitutionsrezept(0).Left + txtSubstitutionsrezept(0).Width + 300
lblSubstitutionsrezept2(0).Top = lblSubstitutionsrezept(0).Top
For i% = 1 To 3
    lblSubstitutionsrezept2(i%).Left = lblSubstitutionsrezept2(i% - 1).Left
    lblSubstitutionsrezept2(i%).Top = lblSubstitutionsrezept(i%).Top
Next i%

With chkSubstitutionsrezept(0)
    .Left = lblSubstitutionsrezept(0).Left
    .Top = lblSubstitutionsrezept(3).Top + lblSubstitutionsrezept(3).Height + 600
End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

'Me.Width = fmeImpAlternativ.Left + fmeImpAlternativ.Width + 2 * wpara.LinksX
Me.Width = lblSubstitutionsrezept2(2).Left + lblSubstitutionsrezept2(2).Width + 2 * wpara.LinksX

With cmdOk
'    .Top = fmeImpAlternativ.Top + fmeImpAlternativ.Height + 150 * wpara.BildFaktor
    .Top = chkSubstitutionsrezept(0).Top + chkSubstitutionsrezept(0).Height + 450
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
        .Top = chkSubstitutionsrezept(0).Top + chkSubstitutionsrezept(0).Height + iAdd + 600
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
        .Enabled = (txtSubstitutionsrezept(0).text <> "") And (txtSubstitutionsrezept(1).text <> "") ' cmdOk.Enabled
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
'                If (c.Name = "chkSubstitutionsrezept") Then
'                    If (c.Index > 0) Then
'                        Load lblchkSubstitutionsrezept(c.Index)
'                    End If
'                    With lblchkSubstitutionsrezept(c.Index)
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

Private Sub txtSubstitutionsrezept_Change(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtSubstitutionsRezept_Change")
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

nlcmdOk.Enabled = (txtSubstitutionsrezept(0).text <> "") And (txtSubstitutionsrezept(1).text <> "")

Call DefErrPop
End Sub

Private Sub txtSubstitutionsRezept_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtSubstitutionsRezept_GotFocus")
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

With txtSubstitutionsrezept(index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

If (index = 0) Then
    iEditModus = 4
Else
    iEditModus = 0
End If

Call DefErrPop
End Sub

'Private Sub txtSubstitutionsRezept_KeyPress(index As Integer, KeyAscii As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("txtSubstitutionsRezept_KeyPress")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
'    Beep
'    KeyAscii = 0
'End If
'
'Call DefErrPop
'End Sub

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






