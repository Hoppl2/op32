VERSION 5.00
Begin VB.Form frmWumsatzAdd 
   ClientHeight    =   3495
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4245
   Begin VB.TextBox txtWumsatzAdd 
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
      Left            =   3000
      MaxLength       =   8
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtWumsatzAdd 
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
      Index           =   2
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtWumsatzAdd 
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
      MaxLength       =   10
      TabIndex        =   1
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
   Begin VB.Label lblWumsatzAdd 
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
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblWumsatzAdd 
      Caption         =   "Umsatz (NICHT rabattfähig)"
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
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblWumsatzAdd 
      Caption         =   "Umsatz (rabattfähig)"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmWumsatzAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "UMSADD.FRM"

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
Dim j%, dat%
Dim h0$, h1$, h2$
Dim WumsatzRec() As WumsatzStruct

h0$ = Trim(txtWumsatzAdd(0).text)
dat% = iDate(h0$)
h1$ = Trim(txtWumsatzAdd(1).text)
h2$ = Trim(txtWumsatzAdd(2).text)

If (dat% = 0) Then
    txtWumsatzAdd(0).SetFocus
    Beep
ElseIf (h1$ = "") And (h2$ = "") Then
    txtWumsatzAdd(1).SetFocus
    Beep
Else
    j% = 1
    ReDim WumsatzRec(j%)
    If (h1$ <> "") Then
        WumsatzRec(j%).Lief = WuFrageErg%
        WumsatzRec(j%).bdatum = h0$
        WumsatzRec(j%).Wert = Val(h1$)
        WumsatzRec(j%).Rabatt = True
        j% = j% + 1
    End If
    If (h2$ <> "") Then
        ReDim Preserve WumsatzRec(j%)
        WumsatzRec(j%).Lief = WuFrageErg%
        WumsatzRec(j%).bdatum = h0$
        WumsatzRec(j%).Wert = Val(h2$)
        WumsatzRec(j%).Rabatt = False
    End If

    Call WumsatzEinzeln(WumsatzRec)
    
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
If (ActiveControl.Name = txtWumsatzAdd(0).Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
        If (ActiveControl.Index = 0) Or (Chr$(KeyAscii) <> "-") Then
            Beep
            KeyAscii = 0
        End If
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
Dim h$, h2$

h$ = "Zusätzlicher Umsatz für "

If (WuFrageErg% > 0) And (WuFrageErg% <= lif.AnzRec) Then
    lif.GetRecord (WuFrageErg% + 1)
    h2$ = RTrim$(lif.kurz)
Else
    h2$ = "??????"
End If
h$ = h$ + h2$
Me.Caption = h$

Call wpara.InitFont(Me)

txtWumsatzAdd(0).Width = TextWidth("99.99.9999") + 90
txtWumsatzAdd(1).Width = txtWumsatzAdd(0).Width
txtWumsatzAdd(2).Width = txtWumsatzAdd(1).Width

h$ = Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000")
h$ = Left$(h$, 4) + Right$(h$, 2)
txtWumsatzAdd(0).text = h$

txtWumsatzAdd(1).text = Space$(8)
txtWumsatzAdd(2).text = txtWumsatzAdd(1).text

txtWumsatzAdd(0).Top = wpara.TitelY%
For i% = 1 To 2
    txtWumsatzAdd(i%).Top = txtWumsatzAdd(i% - 1).Top + txtWumsatzAdd(i% - 1).Height + 90
Next i%

lblWumsatzAdd(0).Left = wpara.LinksX
lblWumsatzAdd(0).Top = txtWumsatzAdd(0).Top
For i% = 1 To 2
    lblWumsatzAdd(i%).Left = lblWumsatzAdd(i% - 1).Left
    lblWumsatzAdd(i%).Top = txtWumsatzAdd(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblWumsatzAdd(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtWumsatzAdd(0).Left = lblWumsatzAdd(0).Left + MaxWi% + 300
For i% = 1 To 2
    txtWumsatzAdd(i%).Left = txtWumsatzAdd(i% - 1).Left
Next i%

cmdOk.Top = lblWumsatzAdd(2).Top + lblWumsatzAdd(2).Height + 150
cmdEsc.Top = cmdOk.Top

MaxWi% = txtWumsatzAdd(0).Left + txtWumsatzAdd(0).Width
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

Call DefErrPop
End Sub

Private Sub txtWumsatzAdd_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtWumsatzAdd_GotFocus")
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
With txtWumsatzAdd(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With
Call DefErrPop
End Sub

