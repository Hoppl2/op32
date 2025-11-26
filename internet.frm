VERSION 5.00
Begin VB.Form frmInternet 
   Caption         =   "Angebote aus dem Internet"
   ClientHeight    =   3555
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4875
   Icon            =   "internet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblInternet 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "INTERNET.FRM"

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdEsc_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub Form_Load()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Load")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

ind% = val(StammdatenPzn$)
Lif1.GetRecord (ind% + 1)
LifZus1.GetRecord (ind% + 1)
            
h2$ = Lif1.kurz
h$ = "Angebote von Lieferant " + h2$ + " einholen" + vbCrLf + vbCrLf + "Internet-Adresse: " + LifZus1.AngebotWWW

With lblInternet
    .Caption = h$
    .Left = wPara1.LinksX
    .Top = wPara1.TitelY
    .Height = TextHeight(h$)
End With

Call wPara1.InitFont(Me)

Font.Bold = False   ' True

lblInternet.Height = TextHeight(h$)

Me.Width = lblInternet.Width + 2 * wPara1.LinksX

cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdEsc.Left = (Me.ScaleWidth - cmdEsc.Width) / 2
cmdEsc.Top = lblInternet.Top + lblInternet.Height + 300

Me.Height = cmdEsc.Top + cmdEsc.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight
Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

Call clsError.DefErrPop
End Sub

