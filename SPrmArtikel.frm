VERSION 5.00
Begin VB.Form frmSPrmArtikel 
   Caption         =   "Sonderprämienartikel hinzufügen"
   ClientHeight    =   3960
   ClientLeft      =   450
   ClientTop       =   2010
   ClientWidth     =   6870
   Icon            =   "SPrmArtikel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   6870
   Begin VB.CommandButton cmdEinzel 
      Caption         =   "F2 Einzelerfassung"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   600
      TabIndex        =   11
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2040
      TabIndex        =   12
      Top             =   2760
      Width           =   1200
   End
   Begin VB.TextBox txtEin 
      Height          =   405
      Index           =   4
      Left            =   4320
      TabIndex        =   9
      Text            =   "1234567890"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtEin 
      Height          =   405
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Text            =   "1234567890"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtEin 
      Height          =   405
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Text            =   "1234567890"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtEin 
      Height          =   405
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Text            =   "1234567890"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtEin 
      Height          =   405
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblEin 
      Caption         =   "bis"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblEin 
      Caption         =   "A&VP von"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblEin 
      Caption         =   "&ATC-Code:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblEin 
      Caption         =   "&Lagercodes:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblEin 
      Caption         =   "&Warengruppen:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSPrmArtikel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "SPRMARTIKEL.FRM"
Private Sub cmdEinzel_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEinzel_Click")
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
Dim s$, pzn As String, h$, txt As String
Dim i%, j%, erg%
Dim satz As Long

s$ = MatchCode(0, pzn, txt, False, False)
Do While Len(s$) > 0
  i% = InStr(s$, vbTab)
  If i% = 0 Then i% = Len(s$)
  j% = InStr(s$, "@")
  Do While j% > 0 And j% < i%
    s$ = Left(s$, j% - 1) + vbTab + Mid(s$, j% + 1)
    j% = InStr(s$, "@")
  Loop
  pzn = Left(s$, 7)
  sp.index = "Unique"
  sp.Seek "=", pzn
  If sp.NoMatch Then
    erg% = ast.IndexSearch(0, pzn, satz)
    If erg% = 0 Then
      sp.AddNew
      sp!pzn = Left(s$, 7)
      sp.Update
      h$ = Left$(s$, i% - 1)
      frmWinPvsOptionen.flxOptionen1(1).AddItem h$
    End If
  End If
  s$ = Mid$(s$, i% + 1)
Loop
Unload Me
Call DefErrPop
End Sub

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
Dim ok As Boolean
Dim i%

ok = False
For i% = 0 To 4
  If Trim(txtEin(i%).Text) > "" Then
    ok = True
    Exit For
  End If
Next i%
If ok Then
  PrgAction = ARTIKEL_EINLESEN

  frmFortschritt.Show vbModal
  With frmWinPvsOptionen.flxOptionen1(1)
    .Row = 1
    .RowSel = 0
    .Col = 1
    .ColSel = 3
    .Sort = 5
  End With
End If
Unload Me
Call DefErrPop

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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
If KeyCode = vbKeyF2 And Shift = 0 Then
  KeyCode = 0
  Call cmdEinzel_Click
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim w As Long
Dim h$, h2$, h3$, FormStr$
Dim c As Object

Call wpara.InitFont(Me)
Me.KeyPreview = True

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.Text) + 90
        c.Text = ""
    End If
Next
On Error GoTo DefErr

lblEin(0).Top = wpara.TitelY
txtEin(0).Top = lblEin(0).Top + (lblEin(i%).Height - txtEin(i%).Height) / 2
txtEin(0).Left = lblEin(0).Left + lblEin(0).Width + 300

For i% = 1 To 3
  lblEin(i%).Left = lblEin(0).Left
  txtEin(i%).Left = txtEin(0).Left
  lblEin(i%).Top = lblEin(i% - 1).Top + lblEin(i% - 1).Height + 300
  txtEin(i%).Top = lblEin(i%).Top + (lblEin(i%).Height - txtEin(i%).Height) / 2
Next i%
lblEin(4).Top = lblEin(3).Top
lblEin(4).Left = txtEin(3).Left + txtEin(3).Width + 300
txtEin(4).Top = txtEin(3).Top
txtEin(4).Left = lblEin(4).Left + lblEin(4).Width + 300

Me.Width = txtEin(4).Left + txtEin(4).Width + 2 * wpara.LinksX

cmdOk.Top = lblEin(4).Top + lblEin(4).Height + 300
cmdEsc.Top = cmdOk.Top
cmdEinzel.Top = cmdOk.Top
cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEinzel.Height = wpara.ButtonY
cmdEinzel.Width = wpara.ButtonX
cmdEinzel.Left = lblEin(0).Left
cmdEsc.Left = Me.ScaleWidth - cmdEsc.Width - 150

cmdOk.Left = Me.Width / 2 - cmdOk.Width / 2

Me.Height = cmdOk.Top + cmdOk.Height + wpara.FrmCaptionHeight + 300

Me.Top = frmWinPvsOptionen.Top + (frmWinPvsOptionen.Height - Me.Height) / 2
Me.Left = frmWinPvsOptionen.Left + (frmWinPvsOptionen.Width - Me.Width) / 2

End Sub


