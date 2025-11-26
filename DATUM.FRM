VERSION 5.00
Begin VB.Form frmDatum 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "PVS - Auswertung"
   ClientHeight    =   4500
   ClientLeft      =   1935
   ClientTop       =   2610
   ClientWidth     =   8745
   Icon            =   "datum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstWer 
      Height          =   3660
      Left            =   4080
      Style           =   1  'Kontrollkästchen
      TabIndex        =   5
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtDatum 
      Height          =   435
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Text            =   "31.1.1999"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtDatum 
      Height          =   435
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Text            =   "1.1.1999"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblMitarbeiter 
      Caption         =   "&Auswertung für folgende Mitarbeiter:"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblUeberleitung 
      Caption         =   "Daten übergeleitet bis 31.12.99"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "&bis:"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Auswertung &von:"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1635
   End
End
Attribute VB_Name = "frmDatum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "DATUM.FRM"
Sub AllesMarkieren()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AllesMarkieren")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error Resume Next

If TypeOf frmDatum.ActiveControl Is TextBox Then
  On Error GoTo DefErr
  frmDatum.ActiveControl.SelStart = 0
  frmDatum.ActiveControl.SelLength = Len(frmDatum.ActiveControl.Text)
End If

Call DefErrPop
End Sub


Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEsc_Click")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

abbruch = True
Unload Me
Call DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Load")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim WoTag%

If PrgAction = VERKAUF_EINLESEN Then
  Me.Caption = "PVS - Überleitung"
  Label1(0).Caption = "Überleitung von"
Else
  Me.Caption = "PVS - Auswertung"
  Label1(0).Caption = "Auswertung von"
End If

Call wpara.InitFont(Me)
Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

cmdOk.Height = wpara.ButtonY
cmdOk.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
lblUeberleitung.Top = wpara.TitelY
lblUeberleitung.Left = wpara.LinksX
Label1(0).Left = lblUeberleitung.Left
Label1(1).Left = lblUeberleitung.Left
Label1(0).Top = lblUeberleitung.Top + lblUeberleitung.Height + 300
Label1(1).Top = Label1(0).Top + Label1(0).Height + 300
txtDatum(0).Top = Label1(0).Top + (Label1(0).Height - txtDatum(0).Height) / 2
txtDatum(1).Top = Label1(1).Top + (Label1(1).Height - txtDatum(1).Height) / 2
lstWer.Top = Label1(0).Top
lstWer.Left = txtDatum(0).Left + txtDatum(0).Width + 300
lblMitarbeiter.Top = wpara.TitelY
lblMitarbeiter.Left = lstWer.Left
lblUeberleitung.Caption = "Daten übergeleitet bis " + Format(UebDatBis, "dd.mm.yy")

If PrgAction = VERKAUF_EINLESEN Then
  lstWer.Visible = False
  lblMitarbeiter.Visible = False
  cmdOk.Top = txtDatum(1).Top + txtDatum(1).Height + 300
  cmdEsc.Top = cmdOk.Top
  Me.Width = txtDatum(0).Left + txtDatum(0).Width + 300
  If ErstesMal% <> 0 Then
    v.GetRecord 2
    UebDatBis = DateAdd("d", -1, dDatum(v.datum))
  End If
  txtDatum(0).Text = Format(DateAdd("d", 1, UebDatBis), "dd.mm.yy")
  
  txtDatum(1).Text = Format(DateAdd("d", -1, Now), "dd.mm.yy")
  WoTag% = WeekDay(Now) - 2
  If WoTag% >= 0 And Val(Format(Now, "hhnn")) >= OffenRec(WoTag%).iBisOffen Then
    txtDatum(1).Text = Format(Now, "dd.mm.yy")
  End If
Else
  lstWer.Visible = True
  lblMitarbeiter.Visible = True
  cmdOk.Top = lstWer.Top + lstWer.Height - cmdOk.Height
  cmdEsc.Top = cmdOk.Top
  Me.Width = lstWer.Left + lstWer.Width + 300
  If sF2Vorschau Then
    txtDatum(0).Text = Format(Now, "dd.mm.yy")
    txtDatum(1).Text = Format(Now, "dd.mm.yy")
  Else
    txtDatum(0).Text = Format(vonAuswD, "dd.mm.yy")
    txtDatum(1).Text = Format(bisAuswD, "dd.mm.yy")
  End If
End If
Me.Height = cmdOk.Top + cmdOk.Height + wpara.FrmCaptionHeight + 300
If PrgAction = VERKAUF_EINLESEN Then    'Setfocus funkt. hier noch nicht -> wird in Gotfocus wieder enabled
  txtDatum(0).Enabled = False
End If
Call DefErrPop
End Sub

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOk_Click")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim Cancel As Boolean
Dim erg As Integer
Dim i%

Call ChecktxtDatum(0, Cancel)

If Cancel Then
  If PrgAction = VERKAUF_EINLESEN Then
    txtDatum(0).Text = Format(DateAdd("d", 1, UebDatBis), "dd.mm.yy")
  Else
    erg = iMsgBox("Bitte überprüfen Sie das eingegebene Datum!", vbOKOnly)
    Call DefErrPop
    Exit Sub
  End If
End If
Call ChecktxtDatum(1, Cancel)
If Cancel Then
  erg = iMsgBox("Bitte überprüfen Sie das eingegebene Datum!", vbOKOnly)
  Call DefErrPop
  Exit Sub
End If

vonAuswD = CDate(txtDatum(0).Text)
bisAuswD = CDate(txtDatum(1).Text)
If PrgAction <> VERKAUF_EINLESEN Then
  If vonAuswD > UebDatBis Then bisAuswD = vonAuswD
  With lstWer
    For i% = 0 To .ListCount - 1
      If .Selected(i%) Then
        Mid$(MitArb$(i% + 1), 21, 1) = "*"
        iMitArb(Val(Mid$(MitArb$(i% + 1), 22, 2))) = True
      Else
        Mid$(MitArb$(i% + 1), 21, 1) = " "
        iMitArb(Val(Mid$(MitArb$(i% + 1), 22, 2))) = False
      End If
    Next i%
  End With
Else

  
End If

abbruch = False

Unload Me
Call DefErrPop

End Sub


Private Sub ChecktxtDatum(Index As Integer, Cancel As Boolean)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ChecktxtDatum")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim erg%
Dim d As Date
Dim s As String

s = txtDatum(Index).Text
s = CheckDatum(s)

If Not IsDate(s) Then
  erg = iMsgBox("Bitte geben Sie ein gültiges Datum ein!", vbOKOnly)
  Cancel = True
Else
  txtDatum(Index).Text = s
  d = CVDate(txtDatum(Index).Text)
  If d > DateValue(Now) Then
    If Year(d) > Year(Now) Or Month(d) > Month(Now) Then
      erg = iMsgBox("Das Datum liegt in der Zukunft!", vbOKOnly)
      Cancel = True
    End If
  End If
  If Index = 1 Then
    If IsDate(txtDatum(0).Text) Then
      If CVDate(txtDatum(0).Text) > d Then
        txtDatum(1).Text = txtDatum(0).Text
        txtDatum(0).Text = Format(d, "ddmmyy")
      End If
    End If
  End If
End If

Call DefErrPop
End Sub


Private Sub txtDatum_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtDatum_GotFocus")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
txtDatum(0).Enabled = True
Call AllesMarkieren
Call DefErrPop
End Sub


Private Sub txtDatum_LostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtDatum_LostFocus")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Cancel As Boolean
Dim d As Date
Dim d1 As Date

'If Index = 0 Then
  Call ChecktxtDatum(Index, Cancel)
  If PrgAction <> VERKAUF_EINLESEN Then
    If Not Cancel Then
      If IsDate(txtDatum(0).Text) And Index = 0 Then
        d = txtDatum(0).Text
        
        If d > UebDatBis Then
          txtDatum(1).Text = Format(d, "dd.mm.yy")
        ElseIf Val(Left(d, 2)) = 1 Then
          d = DateAdd("m", 1, d)
          d = DateAdd("d", -1, d)
        Else
          d1 = d
          Do While Month(d) = Month(d1)
            d = DateAdd("d", 1, d)
          Loop
          d = DateAdd("d", -1, d)
        End If
        txtDatum(1).Text = Format(d, "dd.mm.yy")
      End If
    End If
  End If
'End If
Call DefErrPop
End Sub


