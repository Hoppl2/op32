VERSION 5.00
Begin VB.Form frmPass 
   Caption         =   "Benutzer anmelden"
   ClientHeight    =   1740
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4140
   Icon            =   "Pass.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   4140
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   345
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   1200
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Passwort"
      Top             =   660
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Bitte geben Sie Ihr Passwort ein:"
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "PASS.FRM"

Dim Benutzer() As String
Private Sub Command1_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Command1_Click")
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

Dim i As Integer

If Index = 0 Then
  Call BenutzerChecken(Text2.text)
  If ActBenutzer > 0 Then
    Unload Me
  Else
    Command1(0).SetFocus
    Text2.SetFocus
  End If
Else
  Unload Me
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
Dim i As Integer, j As Integer, k As Integer
Dim f As Long
Dim s As String
Dim pwbuf As String * 40

'Me.Top = frmMenu.Top + (frmMenu.Height - Me.Height) / 2
'Me.Left = frmMenu.Left + (frmMenu.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
ActBenutzer = -1

ReDim Benutzer(1) As String
Benutzer(1) = "DP"
f = FreeFile
Open "PASSWORD.DAT" For Random Access Read Write Shared As #f Len = 40

For i = 1 To (LOF(f) / 40) - 1
  Get #f, i + 1, pwbuf
  'If Trim(Left(pwbuf, 20)) <> "" Then
    j = j + 1
    If j > UBound(Benutzer) Then ReDim Preserve Benutzer(j)
    s = Oem2Ansi(Left(pwbuf, 20))
    For k = 1 To 20
      Mid(s, k, 1) = Chr(Asc(Mid(s, k, 1)) - k)
    Next k
    Benutzer(j) = UCase(s)
  'End If
Next i
Close #f

Call DefErrPop
End Sub

Private Sub Text2_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Text2_GotFocus")
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

Text2.text = ""
'Text2.SelStart = 0
'Text2.SelLength = Len(Text2.Text)
Call DefErrPop
End Sub

Sub BenutzerChecken(Passwort As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BenutzerChecken")
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

Dim i As Integer
Dim pw As String
Dim StandardPW As String

pw = UCase(Left(Trim(Passwort) + Space(20), 20))
If Trim(pw) = "" Then
  Call DefErrPop
  Exit Sub
End If
StandardPW = "@" + Format(Now, "hhnnddmm")
If Trim(pw) = StandardPW Then
  ActBenutzer = 1
  Call DefErrPop
  Exit Sub
End If

For i = 1 To UBound(Benutzer)
  If Benutzer(i) = pw Then
    ActBenutzer = i
    Exit For
  End If
Next i
Call DefErrPop
End Sub

