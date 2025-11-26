VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEditAbholer 
   Caption         =   "Editierung Abholer"
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
   Begin VB.CommandButton cmdF5 
      Caption         =   "&Abholer löschen (F5)"
      Height          =   450
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbholer 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
   End
End
Attribute VB_Name = "frmEditAbholer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Const DefErrModul = "EDITABHOLER.FRM"
'
'Private Sub cmdEsc_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdEsc_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'Unload Me
'
'Call DefErrPop
'End Sub
'
'Private Sub cmdOk_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdOk_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, k%, KistenNr%, erg%
'
'If (ActiveControl.Name = cmdOk.Name) Then
'    Kiste.OpenDatei
'    With flxAbholer
'        For i% = 1 To (.Rows - 1)
'            For j% = 1 To 10
'                If (.TextMatrix(i%, j%) = "") And (.TextMatrix(i%, j% + 10) <> "") Then
'                    KistenNr% = (i% - 1) * 10 + (j% - 1)
'
'                    If (Kiste.Belegt(KistenNr%)) Then
'                        Call Kiste.GetKiste(KistenNr%)
'                        For k% = 0 To 9
'                            erg% = Kiste.ClearInhalt(k%)
'                        Next k%
'                        Kiste.ClearKiste (KistenNr%)
'                    End If
'                End If
'            Next j%
'        Next i%
'    End With
'    Kiste.CloseDatei
'
'    Unload Me
'End If
'
'Call DefErrPop
'End Sub
'
'Private Sub cmdF5_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdF5_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'With flxAbholer
'    .TextMatrix(.row, .col) = ""
'    .SetFocus
'End With
'
'Call DefErrPop
'End Sub
'
'Private Sub flxAbholer_DblClick()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("flxAbholer_RowColChange")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'cmdF5.Value = True
'
'Call DefErrPop
'End Sub
'
'Private Sub flxAbholer_RowColChange()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("flxAbholer_RowColChange")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'With flxAbholer
'    If (.col > 10) Then
'        .col = 10
'    End If
'End With
'
'Call DefErrPop
'End Sub
'
'Private Sub Form_Load()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Form_Load")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%
'Dim wi%
'Dim h$, h2$
'
'Call wpara.InitFont(Me)
'
'With flxAbholer
'    .Cols = 22
'    .Rows = 101
'    .FixedRows = 1
'    .FixedCols = 1
'
'    h$ = ">|"
'    For i% = 0 To 9
'        h$ = h$ + "^" + Format(i%) + "|"
'    Next i%
'    For i% = 0 To 9
'        h$ = h$ + "|"
'    Next i%
'    .FormatString = h$
'
'    For i% = 1 To (.Rows - 1)
'        .TextMatrix(i%, 0) = Format(i% - 1, "0x")
'    Next i%
'
'    .Top = wpara.TitelY
'    .Left = wpara.LinksX
'    .Height = .RowHeight(0) * 21 + 90
'
'    .ColWidth(0) = TextWidth(String(6, "9"))
'    For i% = 1 To 10
'        .ColWidth(i%) = TextWidth(String(6, "9"))
'    Next i%
'    For i% = 11 To 20
'        .ColWidth(i%) = 0
'    Next i%
'    .ColWidth(21) = wpara.FrmScrollHeight
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'
'    Call flxAbholerBefuellen
'
'    .FillStyle = flexFillRepeat
'    .row = 1
'    .col = 1
'    .RowSel = .Rows - 1
'    .ColSel = .Cols - 1
'    .CellFontName = "Symbol"
'    .FillStyle = flexFillSingle
'
'    .row = 1
'    .col = 2
'    .ColSel = .col
'End With
'
'
'Font.Name = wpara.FontName(1)
'Font.Size = wpara.FontSize(1)
'
'Me.Width = flxAbholer.Left + flxAbholer.Width + 2 * wpara.LinksX
'
'With cmdF5
'    .Left = flxAbholer.Left
'    .Top = flxAbholer.Top + flxAbholer.Height + 150 * wpara.BildFaktor
'    .Width = TextWidth(.Caption) + 150
'    .Height = wpara.ButtonY
'End With
'
'With cmdEsc
'    .Top = cmdF5.Top
'    .Width = wpara.ButtonX%
'    .Height = wpara.ButtonY%
'    .Left = flxAbholer.Left + flxAbholer.Width - .Width
'End With
'With cmdOk
'    .Top = cmdF5.Top
'    .Width = wpara.ButtonX%
'    .Height = wpara.ButtonY%
'    .Left = cmdEsc.Left - cmdOk.Width - 150
'End With
'
'Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight
'
'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'If (Me.Left < 0) Then
'    Me.Left = 0
'End If
'
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
'If (Me.Top < 0) Then
'    Me.Top = 0
'End If
'
'Call DefErrPop
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Form_KeyDown")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'If (KeyCode = vbKeyF5) Then
'    cmdF5.Value = True
'End If
'
'Call DefErrPop
'End Sub
'
'Private Sub flxAbholerBefuellen()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("flxAbholerBefuellen")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, KistenNr%
'Dim ch$
'
'Kiste.OpenDatei
'
'With flxAbholer
'    For i% = 0 To 99
'        For j% = 0 To 9
'            ch$ = ""
'            KistenNr% = i% * 10 + j%
'            If (Kiste.Belegt(KistenNr%)) Then
'                ch$ = Chr$(214)
'            ElseIf (Kiste.Gedruckt(KistenNr%)) Then
'                ch$ = Chr$(200)
'            End If
'            .TextMatrix(i% + 1, j% + 1) = ch$
'            .TextMatrix(i% + 1, j% + 11) = ch$
'        Next j%
'    Next i%
'End With
'
'Kiste.CloseDatei
'
'Call DefErrPop
'End Sub
'
'
