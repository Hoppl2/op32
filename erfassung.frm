VERSION 5.00
Begin VB.Form frmErfassung 
   Caption         =   "Manuelle Erfassung"
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
   Begin VB.TextBox txtArtikelName 
      Height          =   285
      Left            =   360
      MaxLength       =   35
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtManuell 
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
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "9999"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtManuell 
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
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "9999"
      Top             =   1110
      Width           =   615
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label lblArtikelName 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblManuell 
      Caption         =   "&Naturalrabatt  (0 - 9999)"
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
      TabIndex        =   3
      Top             =   1890
      Width           =   2415
   End
   Begin VB.Label lblManuell 
      Caption         =   "&Bestellmenge (1 - 9999)"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmErfassung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "ERFASSUNG.FRM"

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
ManuellErg% = False
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
ManuellBm% = Val(txtManuell(0).text)
ManuellNm% = Val(txtManuell(1).text)
ManuellTxt$ = Trim$(txtArtikelName.text)
ManuellErg% = True
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
If (ActiveControl.Name = txtManuell(0).Name) Then
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
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

lblArtikelName.Caption = ManuellTxt$
txtArtikelName.text = ManuellTxt$

Call wpara.InitFont(Me)

lblArtikelName.Top = wpara.TitelY%
lblArtikelName.Left = wpara.LinksX
txtArtikelName.Top = lblArtikelName.Top
txtArtikelName.Left = lblArtikelName.Left

txtManuell(0).Top = lblArtikelName.Top + lblArtikelName.Height + 150
For i% = 1 To 1
    txtManuell(i%).Top = txtManuell(i% - 1).Top + txtManuell(i% - 1).Height + 90
Next i%

lblManuell(0).Left = wpara.LinksX
lblManuell(0).Top = txtManuell(0).Top
For i% = 1 To 1
    lblManuell(i%).Left = lblManuell(i% - 1).Left
    lblManuell(i%).Top = txtManuell(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblManuell(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtManuell(0).Left = lblManuell(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtManuell(i%).Left = txtManuell(i% - 1).Left
Next i%

cmdOk.Top = lblManuell(1).Top + lblManuell(1).Height + 150
cmdEsc.Top = cmdOk.Top

wi1% = lblArtikelName.Left + lblArtikelName.Width
wi2% = txtManuell(0).Left + txtManuell(0).Width
If (wi1% > wi2%) Then
    MaxWi% = wi1%
Else
    MaxWi% = wi2%
End If
Me.Width = MaxWi% + 2 * wpara.LinksX
'Me.Width = txtManuell(0).Left + txtManuell(0).Width + 2 * wpara.LinksX
txtArtikelName.Width = MaxWi%

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.frmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

If (ManuellPzn$ = "9999999") Then
    lblArtikelName.Visible = False
    txtArtikelName.Visible = True
Else
    lblArtikelName.Visible = True
    txtArtikelName.Visible = False
End If

Call DefErrPop
End Sub

Private Sub txtmanuell_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtmanuell_GotFocus")
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
With txtManuell(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With
Call DefErrPop
End Sub

