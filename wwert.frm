VERSION 5.00
Begin VB.Form frmWarenWert 
   Caption         =   "Eingabe Warenwert"
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
   Begin VB.TextBox txtWarenwert 
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
   Begin VB.TextBox txtWarenwert 
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
   Begin VB.Label lblWarenWert 
      Caption         =   " Differenz"
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
   Begin VB.Label lblWarenWert 
      Caption         =   " Warenwert"
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
Attribute VB_Name = "frmWarenWert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "WWERT.FRM"

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

If (txtWarenwert(1).Visible) Then
    lblWarenWert(1).Visible = False
    txtWarenwert(1).Visible = False
    txtWarenwert(0).Enabled = True
    txtWarenwert(0).SetFocus
Else
    Warenwert# = -1#
    Unload Me
End If

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
Dim w#

If (txtWarenwert(1).Visible) Then
    Warenwert# = CDbl(txtWarenwert(0).text)
    Unload Me
Else
    w# = Val(txtWarenwert(0).text)
    If (w# = 0#) Then w# = GesamtWert#
    txtWarenwert(0).text = Format(w#, "0.00")
    txtWarenwert(1).text = Format(GesamtWert# - w#, "0.00")
    lblWarenWert(1).Visible = True
    txtWarenwert(1).Visible = True
    txtWarenwert(0).Enabled = False
    txtWarenwert(1).Enabled = False
End If

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
If (ActiveControl.Name = txtWarenwert(0).Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
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

txtWarenwert(0).Width = TextWidth(String$(10, "9")) + 90
txtWarenwert(1).Width = txtWarenwert(0).Width

txtWarenwert(0).text = ""
txtWarenwert(1).text = ""

txtWarenwert(0).Top = wpara.TitelY%
For i% = 1 To 1
    txtWarenwert(i%).Top = txtWarenwert(i% - 1).Top + txtWarenwert(i% - 1).Height + 90
Next i%

lblWarenWert(0).Left = wpara.LinksX
lblWarenWert(0).Top = txtWarenwert(0).Top
For i% = 1 To 1
    lblWarenWert(i%).Left = lblWarenWert(i% - 1).Left
    lblWarenWert(i%).Top = txtWarenwert(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblWarenWert(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtWarenwert(0).Left = lblWarenWert(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtWarenwert(i%).Left = txtWarenwert(i% - 1).Left
Next i%

cmdOk.Top = lblWarenWert(1).Top + lblWarenWert(1).Height + 150
cmdEsc.Top = cmdOk.Top

MaxWi% = txtWarenwert(0).Left + txtWarenwert(0).Width
wi1% = 2 * wpara.ButtonX + 300
If (wi1% > MaxWi%) Then MaxWi% = wi1%

Me.Width = MaxWi% + 2 * wpara.LinksX
'Me.Width = txtManuell(0).Left + txtManuell(0).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

lblWarenWert(1).Visible = False
txtWarenwert(1).Visible = False

Warenwert# = 0#

Call DefErrPop
End Sub

Private Sub txtWarenwert_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtWarenWert_GotFocus")
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

With txtWarenwert(Index)
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With
Call DefErrPop
End Sub

