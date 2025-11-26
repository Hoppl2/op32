VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   0  'Kein
   ClientHeight    =   3210
   ClientLeft      =   1695
   ClientTop       =   1590
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrEdit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   2160
   End
   Begin MSFlexGridLib.MSFlexGrid flxEdit 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      _Version        =   65541
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
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
      Left            =   0
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

'Editmodus:
'1 ... frei
'2 ... Uhrzeit
'3 ... JN
'4 ... Zahl mit NK
'5 ... EAN13 für Ablaufmonat u. Jahr (?)
'6 ... ABCDV (Taralager)
'9 ... Datum DDMM
'10... Datum komplett

Dim ScanJahr%, ScanMonat%


Private Const DefErrModul = "EDIT.FRM"



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
EditErg% = False
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
Dim i%, ind%, uhr%, st%, min%, falsch%
Dim h2$

If (txtEdit.Visible) Then
    falsch% = False
    EditTxt$ = RTrim(txtEdit.Text)
    If (EditModus% = 2) Then
        uhr% = Val(EditTxt$)
        st% = uhr% \ 100
        min% = uhr% Mod 100
        If (st% < 0) Or (st% >= 24) Or (min% < 0) Or (min% >= 60) Then
            falsch% = True
        End If
    ElseIf EditModus% = 3 Then
      falsch% = False
      EditTxt$ = UCase(Trim(txtEdit.Text))
      If Left$(EditTxt$, 1) = "J" Then
        txtEdit.Text = "J"
      Else
        txtEdit.Text = ""
      End If
    
    ElseIf (EditModus% = 5) Then
        If (EditTxt$ = "") Then
            EditErg% = False
            Unload Me
            Call DefErrPop: Exit Sub
        End If
        If (Len(EditTxt$) > 4) Then
            If (Left$(EditTxt$, 8) = "99999978") Then
                ScanJahr% = Val(Mid$(EditTxt$, 11, 2))
            ElseIf (Left$(EditTxt$, 8) = "99999979") Then
                ScanMonat% = Val(Mid$(EditTxt$, 11, 2))
            End If
            txtEdit.Text = Format(ScanMonat%, "00") + Format(ScanJahr%, "00")
            EditTxt$ = RTrim(txtEdit.Text)
        End If
        
        uhr% = Val(EditTxt$)    'Datum
        st% = uhr% \ 100        'Monat
        min% = uhr% Mod 100     'Jahr
        If (st% <= 0) Or (st% > 12) Or (min% <= 0) Or (min% > 10) Then
            falsch% = True
        End If
    ElseIf EditModus% = 6 Then
        EditTxt$ = UCase(EditTxt$)
        If InStr(EditTxt$, "V") > 0 Then EditTxt$ = "V"
        txtEdit.Text = EditTxt$
    ElseIf (EditModus% = 9) Then
        If (iDate(EditTxt$ + Format(Now, "YY")) = 0) Then falsch% = True
    ElseIf (EditModus% = 10) Then
        If (iDate(EditTxt$) = 0) Then falsch% = True
    End If
    If (falsch%) Then
        Beep
        cmdOk.SetFocus
        txtEdit.SetFocus
        Call DefErrPop: Exit Sub
    End If
ElseIf (lstEdit.Visible) Then
    EditTxt$ = RTrim(lstEdit.Text)
ElseIf (flxEdit.Visible) Then
    With flxEdit
        EditTxt$ = ""
        For i% = 0 To 2
            If (i% >= .Cols) Then Exit For
            If (EditTxt$ <> "") Then EditTxt$ = EditTxt$ + vbTab
            EditTxt$ = EditTxt$ + .TextMatrix(.Row, i%)
        Next i%
'        EditTxt$ = .TextMatrix(.row, 0) + vbTab + .TextMatrix(.row, 1) + vbTab + .TextMatrix(.row, 2)
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
                    If (h2$ = "und") Then
                        EditTxt$ = "und"
                        EditAnzGefunden% = 1
                        EditGef%(0) = 999
                        Exit For
                    Else
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
                End If
            Next i%
        End If
    End With
End If
EditErg% = True
Unload Me
Call DefErrPop
End Sub

Private Sub flxEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEdit_DblClick")
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

cmdOk.Value = True

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
Dim s$

If (ActiveControl.Name = txtEdit.Name) Then
'    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (EditModus% <> 1) Then
      If EditModus% = 3 Then
        If Chr(KeyAscii) <> "J" And Chr$(KeyAscii) <> "j" And KeyAscii <> 8 Then
          KeyAscii = Asc("N")
        End If
      ElseIf EditModus% = 6 Then
        If KeyAscii <> 8 Then
            s$ = Chr(KeyAscii)
            s$ = UCase(s$)
            If InStr("ABCDV", s$) = 0 Then
                Beep
                KeyAscii = 0
            End If
        End If
      Else
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> "-") And ((EditModus% <> 4) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
      End If
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

Call wpara.InitFont(Me)
ScanJahr% = 0
ScanMonat% = 0

cmdOk.Left = 10000
cmdEsc.Left = 10000

Call DefErrPop
End Sub

Private Sub lstEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lstEdit_DblClick")
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
cmdOk.Value = True
Call DefErrPop
End Sub

Private Sub lstMultiEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lstMultiEdit_DblClick")
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
cmdOk.Value = True
Call DefErrPop
End Sub

Private Sub tmrEdit_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtedit_GotFocus")
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

tmrEdit.Enabled = False
cmdOk.Value = True

Call DefErrPop
End Sub

Private Sub txtedit_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtedit_GotFocus")
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
Dim h$

With txtEdit
'    h$ = .Text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .Text = h$
    .SelStart = 0
    .SelLength = Len(.Text)
End With
Call DefErrPop
End Sub

