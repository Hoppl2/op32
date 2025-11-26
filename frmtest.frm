VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   1650
   ClientTop       =   1155
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7440
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   1035
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   720
      ScaleHeight     =   2475
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const DefErrModul = "frmtest.frm"
Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Load")
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
'    Set Picture1.Picture = frmWinWawi.imgToolbar.ListImages(5).Picture
'    Picture1.PaintPicture frmWinWawi.Icon, 900, 100, , , 0, 0, , , vbSrcCopy
'    Picture1.PaintPicture frmWinWawi.imgToolbar.ListImages(5).Picture, 0, 0
    Picture2.Picture = LoadPicture("\wininfo\new.bmp")
    Picture1.PaintPicture Picture2.Picture, 0, 0
    'vbSrcCopy
    Picture1.Refresh
    
Call DefErrPop
End Sub
