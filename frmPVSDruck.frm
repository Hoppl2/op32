VERSION 5.00
Begin VB.Form frmDruck 
   Caption         =   "Personal-Verkaufstatistik"
   ClientHeight    =   1755
   ClientLeft      =   4950
   ClientTop       =   2370
   ClientWidth     =   4455
   Icon            =   "frmPVSDruck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4455
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CheckBox chkDruck 
      Caption         =   "&Summen pro Mitarbeiter drucken"
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   4215
   End
   Begin VB.CheckBox chkDruck 
      Caption         =   "&Gesamtsummen drucken"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   4215
   End
End
Attribute VB_Name = "frmDruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "frmDruck.FRM"

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
EditErg% = False
Unload Me
Call DefErrPop
End Sub

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOK_Click")
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
EditErg% = 0
If chkDruck(0).Value <> 0 Then
  EditErg% = 1
End If
If chkDruck(1).Value <> 0 Then
  EditErg% = EditErg% + 100
End If
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
Dim c As Object

Call wpara.InitFont(Me)

'txtOptionen0(1).Text = String(38, "A")

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.Text) + 90
        c.Text = ""
    End If
Next c
On Error GoTo DefErr

EditErg% = True

chkDruck(0).Top = wpara.TitelY
chkDruck(0).Left = wpara.LinksX
chkDruck(1).Top = chkDruck(0).Top + chkDruck(0).Height + 300
chkDruck(1).Left = wpara.LinksX

cmdOk.Top = chkDruck(1).Top + chkDruck(1).Height + 300
cmdEsc.Top = cmdOk.Top
cmdEsc.Left = chkDruck(1).Left + chkDruck(1).Width - cmdEsc.Width
Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight
Me.Width = chkDruck(1).Left + chkDruck(1).Width + wpara.LinksX
Call DefErrPop
End Sub


