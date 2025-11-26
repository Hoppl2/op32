VERSION 5.00
Begin VB.Form frmEdit2 
   BorderStyle     =   0  'Kein
   ClientHeight    =   2055
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtEdit 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstEdit 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "frmEdit.frm"

Private Sub cmdEsc_Click()
EditErg% = False
Unload Me
End Sub

Private Sub cmdOk_Click()
If (txtEdit.Visible) Then
    EditTxt$ = RTrim(txtEdit.text)
Else
    EditTxt$ = RTrim(lstEdit.text)
End If
EditErg% = True
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (EditModus% = 0) Then
    If (ActiveControl.Name = txtEdit.Name) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Call wPara1.InitFont(Me)
End Sub

Private Sub txtedit_GotFocus()
With txtEdit
    .SelStart = 0
    .SelLength = Len(.text)
End With
End Sub


