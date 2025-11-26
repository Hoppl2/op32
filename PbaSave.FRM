VERSION 5.00
Begin VB.Form frmPbaSave 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Pba-Analyse speichern"
   ClientHeight    =   4500
   ClientLeft      =   1935
   ClientTop       =   2610
   ClientWidth     =   8745
   Icon            =   "PbaSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstGespeicherteAnalysen 
      Height          =   3570
      Left            =   4080
      TabIndex        =   3
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtAnalyseName 
      Height          =   435
      Left            =   2160
      TabIndex        =   1
      Text            =   "GrundAnalyse"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblGespeicherteAnalysen 
      Caption         =   "Bereits &gespeicherte Analysen:"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblAnalyseName 
      Caption         =   "&Name der Analyse:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1635
   End
End
Attribute VB_Name = "frmPbaSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "PBASAVE.FRM"

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
Dim i%

Call wpara.InitFont(Me)

EditErg% = 0

If (PbaWahlModus% = 0) Then
    With lblAnalyseName
        .Top = wpara.TitelY
        .Left = wpara.LinksX
    End With
    
    With txtAnalyseName
        .Top = lblAnalyseName.Top + (lblAnalyseName.Height - .Height) / 2
        .Left = lblAnalyseName.Left + lblAnalyseName.Width + 300
    End With
    
    With lblGespeicherteAnalysen
        .Top = txtAnalyseName.Top + txtAnalyseName.Height + 600
        .Left = lblAnalyseName.Left
    End With
Else
    Me.Caption = "Pba-Analysen vergleichen"
    lblAnalyseName.Visible = False
    txtAnalyseName.Visible = False
    With lblGespeicherteAnalysen
        .Top = wpara.TitelY
        .Left = wpara.LinksX
        .Caption = "Aktuelle Pba-Analyse vergleichen mit:"
        .Width = lstGespeicherteAnalysen.Width
    End With
End If

With lstGespeicherteAnalysen
    .Top = lblGespeicherteAnalysen.Top + lblGespeicherteAnalysen.Height + 75
    .Left = lblAnalyseName.Left

    .Clear
    If (PbaWahlModus% = 1) Then
        .AddItem "Neuer Pba-Analyse"
    End If
    For i% = 0 To (frmPbaDiagramm.cboAnalysen.ListCount - 1)
        .AddItem frmPbaDiagramm.cboAnalysen.List(i%)
    Next i%
    If (PbaWahlModus% = 1) Then
        .ListIndex = 0
    End If
'    If (.ListCount > 0) Then
'        .ListIndex = 0
'    End If
End With

With cmdOk
    .Height = wpara.ButtonY
    .Width = wpara.ButtonX
    .Top = lstGespeicherteAnalysen.Top + lstGespeicherteAnalysen.Height + 300
End With
With cmdEsc
    .Height = wpara.ButtonY
    .Width = wpara.ButtonX
    .Top = cmdOk.Top
End With

Me.Width = lstGespeicherteAnalysen.Width + 2 * wpara.LinksX

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmPbaDiagramm.Left + (frmPbaDiagramm.Width - Me.Width) / 2
Me.Top = frmPbaDiagramm.Top + (frmPbaDiagramm.Height - Me.Height) / 2

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
Dim i%
Dim h$

If (PbaWahlModus% = 0) Then
    If (ActiveControl.Name = lstGespeicherteAnalysen.Name) Then
        h$ = lstGespeicherteAnalysen.Text
    Else
        h$ = txtAnalyseName.Text
    End If
    h$ = Trim(h$)
    
    With frmPbaDiagramm.cboAnalysen
        For i% = 0 To (.ListCount - 1)
            If (UCase(h$) = UCase(.List(i%))) Then
                If (MsgBox("Pba-Analyse '" + h$ + "' überschreiben ?", vbYesNo Or vbDefaultButton2) <> vbYes) Then
                    Call DefErrPop: Exit Sub
                End If
                Exit For
            End If
        Next i%
    End With
Else
    h$ = lstGespeicherteAnalysen.Text
End If

EditErg% = True
EditTxt$ = h$

Unload Me

Call DefErrPop
End Sub

Private Sub txtAnalyseName_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAnalyseName_GotFocus")
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

With txtAnalyseName
    .SelStart = 0
    .SelLength = Len(.Text)
End With

Call DefErrPop
End Sub


