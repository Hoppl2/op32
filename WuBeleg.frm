VERSION 5.00
Begin VB.Form frmWuBeleg 
   Caption         =   "Eingabe Beleg-Daten"
   ClientHeight    =   3495
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4245
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4245
   Begin VB.TextBox txtWuBeleg 
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
      Index           =   1
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtWuBeleg 
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
      Index           =   0
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1110
      Width           =   615
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label lblWuBeleg 
      Caption         =   "Beleg-&Datum"
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
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblWuBeleg 
      Caption         =   "Beleg-&Nummer"
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
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmWuBeleg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "WUBELEG.FRM"

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

If (Trim$(txtWuBeleg(0).text) = "") Then
    txtWuBeleg(0).SetFocus
    Beep
ElseIf (Trim$(txtWuBeleg(1).text) < 6) Then
    txtWuBeleg(1).SetFocus
    Beep
ElseIf (iDate(txtWuBeleg(1).text) = 0) Then
    txtWuBeleg(1).SetFocus
    Beep
Else
    ActBeleg$ = Left$(Trim$(txtWuBeleg(0).text) + Space$(10), 10)
    ActBelegDatum% = iDate(txtWuBeleg(1).text)
    If (RowaAktiv%) Then Call ActProgram.RowaLsDatei
    Unload Me
End If

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyPress")
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
If (ActiveControl.Name = txtWuBeleg(0).Name) Then
    If (ActiveControl.Index = 1) And (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
        Beep
        KeyAscii = 0
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
Dim i%, MaxWi%, wi%, wi1%, wi2%
Dim h$

Call wpara.InitFont(Me)

txtWuBeleg(0).Width = TextWidth(String$(10, "W")) + 90
txtWuBeleg(1).Width = TextWidth("99.99.9999") + 90

txtWuBeleg(0).text = ActBeleg$
If (ActBelegDatum% = 0) Then
    h$ = Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000")
    h$ = Left$(h$, 4) + Right$(h$, 2)
Else
    h$ = sDate(ActBelegDatum%)
End If
txtWuBeleg(1).text = h$

txtWuBeleg(0).Top = wpara.TitelY%
For i% = 1 To 1
    txtWuBeleg(i%).Top = txtWuBeleg(i% - 1).Top + txtWuBeleg(i% - 1).Height + 90
Next i%

lblWuBeleg(0).Left = wpara.LinksX
lblWuBeleg(0).Top = txtWuBeleg(0).Top
For i% = 1 To 1
    lblWuBeleg(i%).Left = lblWuBeleg(i% - 1).Left
    lblWuBeleg(i%).Top = txtWuBeleg(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblWuBeleg(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtWuBeleg(0).Left = lblWuBeleg(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtWuBeleg(i%).Left = txtWuBeleg(i% - 1).Left
Next i%

cmdOk.Top = lblWuBeleg(1).Top + lblWuBeleg(1).Height + 150
cmdEsc.Top = cmdOk.Top

MaxWi% = txtWuBeleg(0).Left + txtWuBeleg(0).Width
Me.Width = MaxWi% + 2 * wpara.LinksX
'Me.Width = txtManuell(0).Left + txtManuell(0).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY

If (BelegModus% = 0) Then
    cmdOk.Left = (Me.Width - cmdOk.Width) / 2
    cmdEsc.Visible = False
Else
    cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
End If

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Private Sub txtWuBeleg_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtWuBeleg_GotFocus")
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
With txtWuBeleg(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With
Call DefErrPop
End Sub

