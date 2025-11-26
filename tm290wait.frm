VERSION 5.00
Begin VB.Form frmTm290Wait 
   Caption         =   "Drucker der TM290-Familie"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4305
   Begin VB.Timer tmrTm290Wait 
      Interval        =   100
      Left            =   1920
      Top             =   1440
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label lblTm290Wait 
      Caption         =   "Bitte das Rezept in den Drucker einlegen ..."
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmTm290Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "TM290WAIT.FRM"

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
Tm290WaitErg% = False
Unload Me
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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%
Dim DRUCKHANDLE%
Dim h$, h2$

Call wpara.InitFont(Me)

With lblTm290Wait
    .Left = wpara.LinksX
    .Top = wpara.TitelY
End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = lblTm290Wait.Left + lblTm290Wait.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = lblTm290Wait.Top + lblTm290Wait.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Function GetPaperEndStatus%(PeStatus$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetPaperEndStatus%")
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
Dim ret%
Dim ch$
  
ret% = True

DruckStr$ = Chr$(27) + "v" + Chr$(1)
Do
    Call SeriellSend(DruckStr$)
    Call SeriellPause
    ch$ = SeriellReceive
    If (ch$ = PeStatus$) Then Exit Do
Loop

DruckStr$ = Chr$(27) + "v" + Chr$(0)
Call SeriellSend(DruckStr$)

Call SeriellPause

With frmAction.comSenden
    .PortOpen = False
    .PortOpen = True
End With

GetPaperEndStatus% = ret%
  
Call DefErrPop
End Function

Private Sub tmrTm290Wait_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrTm290Wait_Timer")
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
Dim ret%

tmrTm290Wait.Enabled = False

ret% = True
If ((Timer - LetztDruckZeit&) < 5) Then
    ret% = GetPaperEndStatus%(Chr$(3))
End If
If (ret%) Then
    ret% = GetPaperEndStatus%(Chr$(0))
End If

Unload Me

Call DefErrPop
End Sub
