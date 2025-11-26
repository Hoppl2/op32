VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   0  'Kein
   ClientHeight    =   3210
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstMultiEdit 
      Height          =   1410
      Left            =   1320
      Style           =   1  'Kontrollkästchen
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtEdit 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstEdit 
      Height          =   1425
      ItemData        =   "frmEdit.frx":0000
      Left            =   0
      List            =   "frmEdit.frx":0002
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmEdit"
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
Dim i%, ind%, uhr%, st%, min%
Dim h2$

If (txtEdit.Visible) Then
    EditTxt$ = RTrim(txtEdit.text)
    If (EditModus% = 2) Then
        uhr% = Val(EditTxt$)
        st% = uhr% \ 100
        min% = uhr% Mod 100
        If (st% < 0) Or (st% >= 24) Or (min% < 0) Or (min% >= 60) Then
            Beep
            cmdOk.SetFocus
            txtEdit.SetFocus
            Exit Sub
        End If
    End If
ElseIf (lstEdit.Visible) Then
    EditTxt$ = RTrim(lstEdit.text)
Else
    EditTxt$ = ""
    EditAnzGefunden% = 0
    With lstMultiEdit
        If (.Selected(0) = False) Then
            For i% = 0 To .ListCount - 1
                If (.Selected(i%)) Then
                    h2$ = LTrim$(RTrim$(.List(i%)))
                    ind% = InStr(h2$, "(")
                    If (ind% > 0) Then
                        h2$ = Mid$(h2$, ind% + 1)
                        ind% = InStr(h2$, ")")
                        h2$ = Left$(h2$, ind% - 1)
                        EditGef%(EditAnzGefunden%) = Val(h2$)
                        EditAnzGefunden% = EditAnzGefunden% + 1
                        EditTxt$ = EditTxt$ + h2$ + ","
                    Else
                        If (Left$(h2$, 1) <> "-") Then
                            h2$ = Mid$(Str$(i%), 2)
                            EditGef%(EditAnzGefunden%) = Val(h2$)
                            EditAnzGefunden% = EditAnzGefunden% + 1
                            EditTxt$ = EditTxt$ + h2$ + ","
                        End If
                    End If
                End If
            Next i%
        End If
    End With
End If
EditErg% = True
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyPress")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (EditModus% = 0) Or (EditModus% = 2) Then
    If (ActiveControl.Name = txtEdit.Name) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
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
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, MaxWi%, wi%, wi1%, wi2%

Call wpara.InitFont(Me)

Call DefErrPop
End Sub

Private Sub lstEdit_DblClick()
cmdOk.Value = True
End Sub

Private Sub lstMultiEdit_DblClick()
cmdOk.Value = True
End Sub

Private Sub txtedit_GotFocus()
With txtEdit
    .SelStart = 0
    .SelLength = Len(.text)
End With
End Sub

