VERSION 5.00
Begin VB.Form frmEditDirekt 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   3210
   ClientLeft      =   1680
   ClientTop       =   1560
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4935
   Begin VB.CheckBox chkEdit 
      BackColor       =   &H80000005&
      Caption         =   "zusätzlich &Fakturen-Rabatt"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtEdit 
      Height          =   495
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "XXXXX"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BackColor       =   &H80000005&
      Caption         =   "&Zeilen-Rabatt:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmEditDirekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "EDITDIREKT.FRM"

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

EditTxt$ = Trim(txtEdit.text)
If (chkEdit.Value) Then EditTxt$ = "-" + EditTxt$
EditErg% = True
Unload Me

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
If (ActiveControl.Name = txtEdit.Name) Then
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
Dim i%, MaxWi%, wi%, wi1%, wi2%, xpos%, ydiff%

Call wpara.InitFont(Me)

lblEdit.Left = wpara.LinksX
lblEdit.Top = wpara.TitelY

xpos% = lblEdit.Left + lblEdit.Width
ydiff% = (txtEdit.Height - lblEdit.Height) / 2

With txtEdit
    .Left = xpos%
    .Top = lblEdit.Top - ydiff%
    .Width = TextWidth(.text) + 90
    .text = ""
End With

With chkEdit
    .Left = wpara.LinksX
    .Top = txtEdit.Top + txtEdit.Height + 300
    .Width = .Width - 150
End With

cmdOk.Left = 10000
cmdEsc.Left = 10000

wi1% = txtEdit.Left + txtEdit.Width
wi2% = chkEdit.Left + chkEdit.Width
If (wi1% < wi2%) Then wi1% = wi2%
Me.Width = wi1% + 150
Me.Height = chkEdit.Top + chkEdit.Height + 150

Call DefErrPop
End Sub

Private Sub txtedit_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtedit_GotFocus")
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

With txtEdit
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

