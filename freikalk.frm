VERSION 5.00
Begin VB.Form frmFreieKalk 
   Caption         =   "Freie Kalkulation"
   ClientHeight    =   5115
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdAMPV 
      Caption         =   "&AMPV"
      Enabled         =   0   'False
      Height          =   450
      Left            =   4800
      TabIndex        =   15
      Top             =   3120
      Width           =   1200
   End
   Begin VB.ComboBox cboPreisKalk 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtPreisKalk 
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
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label lblPreisKalkErg 
      Caption         =   "999999,99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblPreisKalk 
      Caption         =   "= Kalk-Avp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblPreisKalk 
      Caption         =   "+ Auf&schlag (%)"
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
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblPreisKalk 
      Caption         =   "&Preisbasis"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblPreiseWert 
      Caption         =   "999999,99"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblPreiseWert 
      Caption         =   "999999,99"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblPreiseWert 
      Caption         =   "999999,99"
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
      Left            =   2880
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblPreise 
      Caption         =   "Taxe-Aep"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblPreise 
      Caption         =   "Stamm-Aep"
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblPreise 
      Caption         =   "NN-Aep"
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
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frmFreieKalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "FREIKALK.FRM"

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
ManuellTxt$ = lblPreisKalkErg.Caption
ManuellErg% = True
Unload Me
Call DefErrPop
End Sub

Private Sub cmdAMPV_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAMPV_Click")
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
txtPreisKalk.text = "AMPV"
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
If (ActiveControl.Name = txtPreisKalk.Name) Then
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
ManuellTxt$ = Format(0, "0.00")

Call wpara.InitFont(Me)

For i% = 0 To 2
    lblPreiseWert(i%).Caption = Format(FreiKalkPreise#(i%), "0.00")
Next i%

With cboPreisKalk
    .Clear
    .AddItem "NN-Aep"
    .AddItem "Stamm-Aep"
    .AddItem "Taxe-Aep"
    .ListIndex = 0
End With
lblPreisKalkErg.Caption = ""

lblArtikelName.Top = wpara.TitelY%
lblArtikelName.Left = wpara.LinksX


lblPreise(0).Top = lblArtikelName.Top + lblArtikelName.Height + 300
For i% = 1 To 2
    lblPreise(i%).Top = lblPreise(i% - 1).Top + lblPreise(i% - 1).Height + 90
Next i%
For i% = 0 To 2
    lblPreiseWert(i%).Top = lblPreise(i%).Top
Next i%

For i% = 0 To 2
    lblPreise(i%).Left = wpara.TitelY%
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblPreise(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblPreiseWert(0).Left = lblPreise(0).Left + MaxWi% + 600
For i% = 1 To 2
    lblPreiseWert(i%).Left = lblPreiseWert(i% - 1).Left
Next i%



lblPreisKalk(0).Top = lblPreiseWert(2).Top + lblPreiseWert(2).Height + 450
For i% = 1 To 2
    lblPreisKalk(i%).Top = lblPreisKalk(i% - 1).Top + cboPreisKalk.Height + 150
Next i%
cboPreisKalk.Top = lblPreisKalk(0).Top
txtPreisKalk.Top = lblPreisKalk(1).Top
lblPreisKalkErg.Top = lblPreisKalk(2).Top

For i% = 0 To 2
    lblPreisKalk(i%).Left = wpara.TitelY%
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblPreisKalk(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

cboPreisKalk.Left = lblPreisKalk(0).Left + MaxWi% + 600
txtPreisKalk.Left = cboPreisKalk.Left
lblPreisKalkErg.Left = cboPreisKalk.Left

With cmdAMPV
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Top = txtPreisKalk.Top
    .Left = cboPreisKalk.Left + cboPreisKalk.Width + 300
End With


cmdOk.Top = lblPreisKalkErg.Top + lblPreisKalkErg.Height + 300
cmdEsc.Top = cmdOk.Top

wi1% = lblArtikelName.Left + lblArtikelName.Width
wi2% = lblPreiseWert(0).Left + lblPreiseWert(0).Width
If (wi1% > wi2%) Then
    MaxWi% = wi1%
Else
    MaxWi% = wi2%
End If
wi1% = cmdAMPV.Left + cmdAMPV.Width
If (wi1% > MaxWi%) Then
    MaxWi% = wi1%
End If

Me.Width = MaxWi% + 2 * wpara.LinksX
'Me.Width = txtManuell(0).Left + txtManuell(0).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.frmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call AvpKalkulieren

Call DefErrPop
End Sub

Private Sub txtPreisKalk_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtPreisKalk_GotFocus")
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
With txtPreisKalk
    .SelStart = 0
    .SelLength = Len(.text)
End With
cmdAMPV.Enabled = True
Call DefErrPop
End Sub

Private Sub txtPreisKalk_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtPreisKalk_LostFocus")
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
cmdAMPV.Enabled = False
Call DefErrPop
End Sub

Private Sub txtPreisKalk_Change()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtPreisKalk_Change")
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
Call AvpKalkulieren
Call DefErrPop
End Sub

Private Sub cboPreisKalk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboPreisKalk_Click")
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
Call AvpKalkulieren
Call DefErrPop
End Sub

Private Sub AvpKalkulieren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AvpKalkulieren")
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
Dim aufschl%, mw%
Dim PreisBasis#, KalkAvp#

PreisBasis# = FreiKalkPreise#(cboPreisKalk.ListIndex)
If (txtPreisKalk.text = "AMPV") Then
    aufschl% = 211
Else
    aufschl% = Val(txtPreisKalk.text)
End If
mw% = para.mwst(Val(FreiKalkMw$))
        
If (aufschl% = 211) Then
    KalkAvp# = CalcAMPV(PreisBasis#, mw%)
Else
    KalkAvp# = PreisBasis# * (100# + aufschl%) / 100#
    KalkAvp# = FNX(KalkAvp# * (100# + mw%) / 100#)
End If

lblPreisKalkErg.Caption = Format(KalkAvp#, "0.00")

Call DefErrPop
End Sub

