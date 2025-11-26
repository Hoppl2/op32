VERSION 5.00
Begin VB.Form frmMdb2Sql 
   Caption         =   "Konvertieren der DOS-Artikeldaten"
   ClientHeight    =   4380
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   6795
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Fernsteuerung"
   ScaleHeight     =   4380
   ScaleWidth      =   6795
   Visible         =   0   'False
   Begin VB.ListBox lstMerkzettel 
      Height          =   300
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picStatus 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   600
      ScaleHeight     =   2295
      ScaleWidth      =   5055
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      Begin VB.PictureBox picProgress 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   945
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Width           =   1200
   End
End
Attribute VB_Name = "frmMdb2Sql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HochfahrenAktiv%

Private Const DefErrModul = "MDB2SQL.FRM"

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

'ActionAbbruch% = True
KonvAbbruch = True

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
Dim i%

Call wpara.InitFont(Me)

KonvAbbruch = 0

Font.Bold = False   ' True

With picStatus
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    .Width = TextWidth(String(60, "X"))
    .Height = TextHeight("Äg") * 14
'    .Left = (Me.ScaleWidth - .Width) / 2
'    .Top = fmeAction.Top + fmeAction.Height + 300
    .FontBold = True
    OrgFontName = .FontName
End With
Me.Width = picStatus.Left + picStatus.Width + 2 * wpara.LinksX + 120

With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.ScaleWidth - .Width) / 2
    .Top = picStatus.Top + picStatus.Height + 150
End With
Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + 600 ' FrmMenuHeight + FrmCaptionHeight

Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

HochfahrenAktiv% = True

Call DefErrPop
End Sub

