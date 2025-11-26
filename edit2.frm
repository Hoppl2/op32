VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdit2 
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
   Begin MSFlexGridLib.MSFlexGrid flxEdit 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.ListBox lstMultiEdit 
      Height          =   1410
      Left            =   1320
      Style           =   1  'Kontrollkästchen
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3600
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtEdit 
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstEdit 
      Height          =   1425
      ItemData        =   "edit2.frx":0000
      Left            =   0
      List            =   "edit2.frx":0002
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

Dim ScanJahr%, ScanMonat%


Private Const DefErrModul = "EDIT2.FRM"

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
EditErg% = False
Unload Me
Call clsError.DefErrPop
End Sub

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdOk_Click")
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
Dim i%, ind%, uhr%, st%, min%, falsch%
Dim h2$

If (txtEdit.Visible) Then
    falsch% = False
    EditTxt$ = RTrim(txtEdit.text)
    If (EditModus% = 2) Then
        uhr% = Val(EditTxt$)
        st% = uhr% \ 100
        min% = uhr% Mod 100
        If (st% < 0) Or (st% >= 24) Or (min% < 0) Or (min% >= 60) Then
            falsch% = True
        End If
    ElseIf (EditModus% = 5) Then
        If (Len(EditTxt$) > 4) Then
            If (Left$(EditTxt$, 8) = "99999978") Then
                ScanJahr% = Val(Mid$(EditTxt$, 11, 2))
            ElseIf (Left$(EditTxt$, 8) = "99999979") Then
                ScanMonat% = Val(Mid$(EditTxt$, 11, 2))
            End If
            txtEdit.text = Format(ScanMonat%, "00") + Format(ScanJahr%, "00")
            EditTxt$ = RTrim(txtEdit.text)
        End If
        
        uhr% = Val(EditTxt$)    'Datum
        st% = uhr% \ 100        'Monat
        min% = uhr% Mod 100     'Jahr
        If (st% <= 0) Or (st% > 12) Or (min% <= 0) Or (min% > 20) Then
            falsch% = True
        End If
    ElseIf (EditModus% = 6) Then
        falsch% = True
        If (Len(EditTxt$) = 6) Then
            If (clsOpTool.iDate(EditTxt$) <> 0) Then falsch% = False
        End If
    End If
    If (falsch%) Then
        Beep
        cmdOk.SetFocus
        txtEdit.SetFocus
        Call clsError.DefErrPop: Exit Sub
    End If
ElseIf (lstEdit.Visible) Then
    EditTxt$ = RTrim(lstEdit.text)
ElseIf (flxEdit.Visible) Then
    With flxEdit
        EditTxt$ = .TextMatrix(.row, 0) + vbTab + .TextMatrix(.row, 1) + vbTab + .TextMatrix(.row, 2) + vbTab + .TextMatrix(.row, 3)
    End With
Else
    EditTxt$ = ""
    EditAnzGefunden% = 0
    EditGef%(0) = 0
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
Call clsError.DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyPress")
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
If (ActiveControl.Name = txtEdit.Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (EditModus% <> 1) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((EditModus% <> 4) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If
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
Dim i%, MaxWi%, wi%, wi1%, wi2%

Call wPara1.InitFont(Me)
ScanJahr% = 0
ScanMonat% = 0

cmdOk.Left = 10000
cmdEsc.Left = 10000

Call clsError.DefErrPop
End Sub

Private Sub lstEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lstEdit_DblClick")
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
cmdOk.Value = True
Call clsError.DefErrPop
End Sub

Private Sub lstMultiEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("lstMultiEdit_DblClick")
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
cmdOk.Value = True
Call clsError.DefErrPop
End Sub

Private Sub txtedit_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtedit_GotFocus")
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
Dim i%
Dim h$

With txtEdit
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With
Call clsError.DefErrPop
End Sub

