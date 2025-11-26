VERSION 5.00
Begin VB.Form frmWuAlternativ 
   Caption         =   "Alternativ-Artikel"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4305
   Begin VB.CheckBox chkWuAlternativ 
      Caption         =   "Alte &Stammdaten löschen"
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CheckBox chkWuAlternativ 
      Caption         =   "Strich&code übernehmen"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CheckBox chkWuAlternativ 
      Caption         =   "&Preise übernehmen"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2520
      TabIndex        =   4
      Top             =   3600
      Width           =   1200
   End
End
Attribute VB_Name = "frmWuAlternativ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WUALTERNATIV.FRM"

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
WuFrageErg% = 0
If (chkWuAlternativ(0).Value = 1) Then WuFrageErg% = WuFrageErg% Or 1
If (chkWuAlternativ(1).Value = 1) Then WuFrageErg% = WuFrageErg% Or 2
If (chkWuAlternativ(2).Value = 1) Then WuFrageErg% = WuFrageErg% Or 4
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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%
Dim h$, h2$

WuFrageErg% = -1

Call wpara.InitFont(Me)

For i% = 0 To 2
    chkWuAlternativ(i%).Left = wpara.LinksX
    If (i% = 0) Then
        chkWuAlternativ(i%).Top = wpara.TitelY
    Else
        chkWuAlternativ(i%).Top = chkWuAlternativ(i% - 1).Top + chkWuAlternativ(i% - 1).Height + 90
    End If
Next i%

With cmdOk
    .Top = chkWuAlternativ(2).Top + chkWuAlternativ(2).Height + 300 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
End With
With cmdEsc
    .Top = cmdOk.Top
    .Width = cmdOk.Width
    .Height = cmdOk.Height
End With

Me.Width = 2 * cmdEsc.Width + 900

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

