VERSION 5.00
Begin VB.Form frmAbbruch 
   Caption         =   "Rufzeitentabelle"
   ClientHeight    =   3360
   ClientLeft      =   720
   ClientTop       =   2640
   ClientWidth     =   7800
   ControlBox      =   0   'False
   Icon            =   "abbruch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7800
   Begin VB.CommandButton cmdBlind 
      Caption         =   "Blind-Bestellung"
      Height          =   450
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdWarten 
      Caption         =   "Weiter warten"
      Default         =   -1  'True
      Height          =   450
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   2040
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "Abbruch"
      Height          =   450
      Left            =   5160
      TabIndex        =   3
      Top             =   2400
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aktiv senden"
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label lblAbbruch 
      Caption         =   "Lieferant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "frmAbbruch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const DefErrModul = "ABBRUCH.FRM"

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdok_Click")
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
AutomatikFertig% = 0
Unload Me
Call DefErrPop
End Sub

Private Sub cmdwarten_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdWarten_Click")
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
AutomatikFertig% = 1
Unload Me
Call DefErrPop
End Sub

Private Sub cmdblind_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdBlind_Click")
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
AutomatikFertig% = 3
Unload Me
Call DefErrPop
End Sub

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdesc_Click")
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
AutomatikFertig% = 2
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
Dim breite%

Caption = "Sendeauftrag für " + LiefName1$

lblAbbruch.Caption = String(30, "X")
Call wpara.InitFont(Me)
lblAbbruch.Caption = "Problem: " + AutomatikFehler$

lblAbbruch.Left = wpara.LinksX%
lblAbbruch.Top = wpara.TitelY%
lblAbbruch.Height = 3 * TextHeight("Äg") + 90

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdOk.Top = lblAbbruch.Top + lblAbbruch.Height + 150
cmdWarten.Top = cmdOk.Top
cmdBlind.Top = cmdOk.Top
cmdEsc.Top = cmdOk.Top

cmdOk.Width = TextWidth(cmdOk.Caption) + 150
cmdOk.Height = wpara.ButtonY%
cmdWarten.Width = TextWidth(cmdWarten.Caption) + 150
cmdWarten.Height = wpara.ButtonY%
cmdBlind.Width = TextWidth(cmdBlind.Caption) + 150
cmdBlind.Height = wpara.ButtonY%
cmdEsc.Width = wpara.ButtonX%
cmdEsc.Height = wpara.ButtonY%

breite% = cmdOk.Width + cmdWarten.Width + cmdBlind.Width + cmdEsc.Width + 450
If (breite% > lblAbbruch.Width) Then
    lblAbbruch.Width = breite%
End If

Me.Width = lblAbbruch.Left + lblAbbruch.Width + 2 * wpara.LinksX%

cmdWarten.Left = (ScaleWidth - breite%) / 2
cmdOk.Left = cmdWarten.Left + cmdWarten.Width + 150
cmdBlind.Left = cmdOk.Left + cmdOk.Width + 150
cmdEsc.Left = cmdBlind.Left + cmdBlind.Width + 150

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
frmAction.WindowState = vbMinimized

If (Dir("wwfehler.wav") <> "") Then
    Call PlaySound("\user\wwfehler.wav", 0, SND_FILENAME Or SND_ASYNC)
End If
    
tmrFocus.Enabled = True

Call DefErrPop
End Sub

Private Sub Form_Paint()

Call FensterImVordergrund

End Sub

Private Sub tmrFocus_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrFocus_Timer")
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

tmrFocus.Enabled = False
cmdWarten.SetFocus

Call DefErrPop
End Sub

Sub FensterImVordergrund()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FensterImVordergrund")
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

If (Me.WindowState = vbMinimized) Or (Me.WindowState = vbMaximized) Then Me.WindowState = vbNormal
Me.SetFocus
DoEvents
Call SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
'Call SetForegroundWindow(Me.hWnd)
Call DefErrPop
End Sub


