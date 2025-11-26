VERSION 5.00
Begin VB.Form frmBestellung 
   Caption         =   "Bestellung"
   ClientHeight    =   1830
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4245
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4245
   Begin VB.TextBox txtBestellung 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   0
      Top             =   390
      Width           =   735
   End
   Begin VB.CommandButton cmdBestellungEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdBestellungOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label lblBestellung 
      Caption         =   "Bestellmenge (1 - 9999)"
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
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmBestellung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DefErrModul = "FRMBES~2.FRM"
Private Sub cmdBestellungEsc_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdBestellungEsc_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Unload Me
Call DefErrPop
End Sub

Private Sub cmdBestellungOk_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdBestellungOk_Click")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call SpeicherBestellung
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
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
    Beep
    KeyAscii = 0
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
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
Call DefErrPop
End Sub

Sub SpeicherBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherBestellung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim bMenge%, BESTZUSA%, AnzDS%
Dim bSatz$

bMenge% = val(txtBestellung.text)
If (bMenge% > 0) Then
    BESTZUSA% = FileOpen%("bestzusa.dat", "RW")
    If (BESTZUSA% > 0) Then
        If (LOF(BESTZUSA%) = 0) Then
            bSatz$ = String(11, 0)
            Put #BESTZUSA%, , bSatz$
            Seek #BESTZUSA%, 1
        End If
        Get #BESTZUSA%, , AnzDS%
        AnzDS% = AnzDS% + 1
        Seek #BESTZUSA%, AnzDS * 11& + 1
        bSatz$ = Format(TaxeRec!pzn, "0000000") + Format(bMenge%, "0000")
        Put #BESTZUSA%, , bSatz$
        Seek #BESTZUSA%, 1
        Put #BESTZUSA%, , AnzDS%
        Close #BESTZUSA%
    End If
End If

Call DefErrPop
End Sub
