VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmZusatz 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   8460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8460
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   4080
      Picture         =   "Zusatz.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "Zusatz.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4560
      Picture         =   "Zusatz.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Edit (F2)"
      Height          =   450
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1200
   End
   Begin VB.TextBox txtZusatz 
      BorderStyle     =   0  'Kein
      Height          =   255
      Index           =   0
      Left            =   0
      MaxLength       =   37
      TabIndex        =   2
      Tag             =   "0"
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxZusatz 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmZusatz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "ZUSATZ.FRM"

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

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF2_Click")
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

'flxZusatz.Visible = False
'For i% = 0 To 4
'    txtZusatz(i%).Visible = True
'Next i%
Call ZeigeTextBoxen

With flxZusatz
'        .BackColor = vbWhite
'        .BackColorBkg = vbWhite
    .HighLight = flexHighlightNever
    .TopRow = .FixedRows
    .row = .FixedRows
    .col = 0
    .ColSel = .Cols - 1
'    txtZusatz(0).text = .TextMatrix(1, 0)
End With
With txtZusatz(0)
    If (ZusatzFensterTyp$ = ZUSATZ_ARTIKEL) Then
        .MaxLength = 37
    Else
        .MaxLength = 60
    End If
    .text = flxZusatz.TextMatrix(1, 0)
    .SelStart = 0
    .SelLength = Len(.text)
End With

If (iNewLine) Then
    nlcmdF2.Enabled = False
    
    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300
    
    nlcmdOk.Cancel = False
    
    nlcmdEsc.Cancel = True
    nlcmdEsc.Visible = True
Else
    cmdF2.Enabled = False
    
    cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
    
    cmdOk.Cancel = False
    
    cmdEsc.Cancel = True
    cmdEsc.Visible = True
End If

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
Dim ind%

'If (txtZusatz(0).Visible) Then
'    If (ActiveControl.Name = txtZusatz(0).Name) Then
'        ind% = ActiveControl.index
'        If (Trim(ActiveControl.text) = "") Or (ind% >= (ZusatzAnzTxt% - 1)) Then
'            If (iNewLine) Then
'                nlcmdOk.SetFocus
'            Else
'                cmdOk.SetFocus
'            End If
'        Else
'            txtZusatz(ind% + 1).SetFocus
'        End If
'    Else
'        Call clsDialog.ZusatzSpeichern
'        Unload Me
'    End If
'Else
'    Unload Me
'End If

If (txtZusatz(0).Visible) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
'        ind% = ActiveControl.index
        ind% = flxZusatz.row - 1
                
        With flxZusatz
            .TextMatrix(.row, 0) = txtZusatz(0).text
        
'            If (Trim(ActiveControl.text) = "") Or ((ZusatzFensterTyp$ <> ZUSATZ_KUNDEN) And (ind% >= (ZusatzAnzTxt% - 1))) Then
            If (Trim(ActiveControl.text) = "") Or ((ind% >= (ZusatzAnzTxt% - 1))) Then
                If (iNewLine) Then
                    nlcmdOk.SetFocus
                Else
                    cmdOk.SetFocus
                End If
'            ElseIf (ZusatzFensterTyp$ = ZUSATZ_KUNDEN) Then
'                If (.row >= (.Rows - 1)) Then
'                    .Rows = .Rows + 1
'                End If
'                .row = .row + 1
'                If (.row > (.TopRow + 19)) Then
'                    .TopRow = .TopRow + 1
'                End If
'                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
'                If (iNewLine = 0) Then
'                    txtZusatz(0).Top = txtZusatz(0).Top + 45
'                End If
'                txtZusatz(0).text = .TextMatrix(.row, 0)
            ElseIf (.row < (.Rows - 1)) Then
                .row = .row + 1
                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
                If (iNewLine = 0) Then
                    txtZusatz(0).Top = txtZusatz(0).Top + 45
                End If
                txtZusatz(0).text = .TextMatrix(.row, 0)
            End If
        End With
        With txtZusatz(0)
            .SelStart = 0
            .SelLength = Len(.text)
        End With
    Else
        If (ZusatzFensterTyp$ = ZUSATZ_ARTIKEL) Then
            If (ArtikelDBok) Then
                Call ArtikelDB1.ZusatzSpeichern(ZusatzPzn$)
            Else
                Call Zus1.ZusatzSpeichern(ZusatzPzn$)
            End If
        ElseIf (ZusatzFensterTyp$ = ZUSATZ_LIEFERANTEN) Then
            Call lZus1.ZusatzSpeichern(ZusatzPzn$)
'        ElseIf (ZusatzFensterTyp$ = ZUSATZ_KUNDEN) Then
'            Call kZus1.ZusatzSpeichern(ZusatzPzn$)
        End If
        Unload Me
    End If
Else
    Unload Me
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyDown")
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
Dim ind%

'If (iNewLine) Then
'    If (KeyCode = vbKeyF2) Then
'        nlcmdF2.Value = True
'    ElseIf (KeyCode = vbKeyDown) Then
'        If (ActiveControl.Name = txtZusatz(0).Name) Then
'            ind% = ActiveControl.index
'            If (ind% < (ZusatzAnzTxt% - 1)) Then
'                txtZusatz(ind% + 1).SetFocus
'            End If
'            KeyCode = 0
'        End If
'    ElseIf (KeyCode = vbKeyUp) Then
'        If (ActiveControl.Name = txtZusatz(0).Name) Then
'            ind% = ActiveControl.index
'            If (ind% > 0) Then
'                txtZusatz(ind% - 1).SetFocus
'            End If
'            KeyCode = 0
'        End If
'    End If
'Else
'    If (KeyCode = vbKeyF2) Then
'        cmdF2.Value = True
'    ElseIf (KeyCode = vbKeyDown) Then
'        If (ActiveControl.Name = txtZusatz(0).Name) Then
'            ind% = ActiveControl.index
'            If (ind% < (ZusatzAnzTxt% - 1)) Then
'                txtZusatz(ind% + 1).SetFocus
'            End If
'            KeyCode = 0
'        End If
'    ElseIf (KeyCode = vbKeyUp) Then
'        If (ActiveControl.Name = txtZusatz(0).Name) Then
'            ind% = ActiveControl.index
'            If (ind% > 0) Then
'                txtZusatz(ind% - 1).SetFocus
'            End If
'            KeyCode = 0
'        End If
'    End If
'End If

If (KeyCode = vbKeyF2) Then
    If (iNewLine) Then
        nlcmdF2.Value = True
    Else
        cmdF2.Value = True
    End If
ElseIf (KeyCode = vbKeyDown) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        With flxZusatz
            .TextMatrix(.row, 0) = txtZusatz(0).text
'            If (ZusatzFensterTyp$ = ZUSATZ_KUNDEN) Then
'                If (.row >= (.Rows - 1)) Then
'                    .Rows = .Rows + 1
'                End If
'                .row = .row + 1
'                If (.row > (.TopRow + 19)) Then
'                    .TopRow = .TopRow + 1
'                End If
'                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
'                If (iNewLine = 0) Then
'                    txtZusatz(0).Top = txtZusatz(0).Top + 45
'                End If
'                txtZusatz(0).text = .TextMatrix(.row, 0)
'            ElseIf (.row < (.Rows - 1)) Then
            If (.row < (.Rows - 1)) Then
                .row = .row + 1
                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
                If (iNewLine = 0) Then
                    txtZusatz(0).Top = txtZusatz(0).Top + 45
                End If
                txtZusatz(0).text = .TextMatrix(.row, 0)
            End If
        End With
        With txtZusatz(0)
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        KeyCode = 0
    End If
ElseIf (KeyCode = vbKeyUp) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        With flxZusatz
            .TextMatrix(.row, 0) = txtZusatz(0).text
'            If (ZusatzFensterTyp$ = ZUSATZ_KUNDEN) Then
'                .TextMatrix(.row, 0) = txtZusatz(0).text
'                If (.row > .FixedRows) Then
'                    .row = .row - 1
'                End If
'                If (.row < .TopRow) Then
'                    .TopRow = .row
'                End If
'                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
'                If (iNewLine = 0) Then
'                    txtZusatz(0).Top = txtZusatz(0).Top + 45
'                End If
'                txtZusatz(0).text = .TextMatrix(.row, 0)
'            ElseIf (.row > .FixedRows) Then
            If (.row > .FixedRows) Then
                .row = .row - 1
                txtZusatz(0).Top = .Top + (.row - .TopRow + 1) * .RowHeight(1)
                If (iNewLine = 0) Then
                    txtZusatz(0).Top = txtZusatz(0).Top + 45
                End If
                txtZusatz(0).text = .TextMatrix(.row, 0)
            End If
        End With
        With txtZusatz(0)
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        KeyCode = 0
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iAdd%, iAdd2%, x%, y%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)


Font.Bold = False   ' True

With flxZusatz
    .Cols = 2
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    
    .FormatString = "<Zusatztext"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(1) = 0
    If (ZusatzFensterTyp$ = ZUSATZ_ARTIKEL) Then
        .ColWidth(0) = TextWidth(String(26, "W"))
        ZusatzAnzTxt% = 5
    ElseIf (ZusatzFensterTyp$ = ZUSATZ_LIEFERANTEN) Then
        .ColWidth(0) = TextWidth(String(40, "W"))
        ZusatzAnzTxt% = 10
'    ElseIf (ZusatzFensterTyp$ = ZUSATZ_KUNDEN) Then
'        .ColWidth(0) = TextWidth(String(40, "W"))
'        .ColWidth(1) = wPara1.FrmScrollHeight
'        ZusatzAnzTxt% = 20
    End If
    
    .Rows = ZusatzAnzTxt% + 1
    .Height = .RowHeight(0) * .Rows + 90
    .Width = .ColWidth(0) + .ColWidth(1) + 90
End With

'For i% = 1 To (ZusatzAnzTxt% - 1)
'    Load txtZusatz(i%)
'    txtZusatz(i%).TabIndex = i% + 2
'Next i%

'For i% = 0 To (ZusatzAnzTxt% - 1)
'    With txtZusatz(i%)
''        .Top = flxZvusatz.Top + i * flxZusatz.RowHeight(1)
'        .Left = flxZusatz.Left + 45
'        .Height = flxZusatz.RowHeight(1)
'        .Width = flxZusatz.Width - 90
'    End With
'Next i%

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

cmdF2.Width = TextWidth(cmdF2.Caption) + 150
cmdF2.Height = wPara1.ButtonY
cmdF2.Left = flxZusatz.Left + flxZusatz.Width + 150
cmdF2.Top = flxZusatz.Top

Me.Width = cmdF2.Left + cmdF2.Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
cmdOk.Top = flxZusatz.Top + flxZusatz.Height + 150

cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdEsc.Top = cmdOk.Top

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

flxZusatz.Visible = True
'For i% = 0 To (ZusatzAnzTxt% - 1)
'    txtZusatz(i%).Visible = False
'Next i%
txtZusatz(0).Visible = False

cmdOk.default = True
cmdOk.Cancel = True
cmdOk.Visible = True

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxZusatz
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wPara1.nlFlexBackColor    'vbWhite
        .BackColorBkg = wPara1.nlFlexBackColor    'vbWhite
        .BackColorFixed = wPara1.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wPara1.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    cmdF2.Left = cmdF2.Left + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxZusatz.Top = flxZusatz.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    cmdF2.Top = cmdF2.Top + iAdd2
    
'    For i% = 0 To (ZusatzAnzTxt% - 1)
'        With txtZusatz(i%)
'            .Left = flxZusatz.Left + 45
'        End With
'    Next i%

    Height = Height + iAdd2

    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = flxZusatz.Top + flxZusatz.Height + 600 * iFaktorY
        .Top = .Top + iAdd
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .default = cmdOk.default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
        .Left = nlcmdOk.Left + .Width + 300
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdF2
        .Init
        .Left = cmdF2.Left
        .Top = cmdF2.Top
        .Caption = cmdF2.Caption
        .TabIndex = cmdF2.TabIndex
        .Enabled = cmdF2.Enabled
        .Visible = True 'cmdF2.Visible
        .AutoSize = True
    End With
    cmdF2.Visible = False

    Me.Width = nlcmdF2.Left + nlcmdF2.Width + 600 * iFaktorX
    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top)
'    RoundRect hdc, (flxZusatz.Left - iAdd) / Screen.TwipsPerPixelX, (flxZusatz.Top - iAdd) / Screen.TwipsPerPixelY, (flxZusatz.Left + flxZusatz.Width + iAdd) / Screen.TwipsPerPixelX, (flxZusatz.Top + flxZusatz.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    
    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
End If



Call clsError.DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Paint")
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
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (Para1.Newline) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top, False)
    RoundRect hdc, (flxZusatz.Left - iAdd) / Screen.TwipsPerPixelX, (flxZusatz.Top - iAdd) / Screen.TwipsPerPixelY, (flxZusatz.Left + flxZusatz.Width + iAdd) / Screen.TwipsPerPixelX, (flxZusatz.Top + flxZusatz.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Sub ZeigeTextBoxen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeTextBoxen")
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

'For i% = 0 To (ZusatzAnzTxt% - 1)
'    With txtZusatz(i%)
'        .Top = flxZusatz.Top + (i + 1) * flxZusatz.RowHeight(1) + 45
'        .Visible = True
'        .ZOrder 0
'    End With
'Next i%
'txtZusatz(0).SetFocus

With txtZusatz(0)
    .Top = flxZusatz.Top + flxZusatz.RowHeight(1)
    .Left = flxZusatz.Left
    If (iNewLine = 0) Then
        .Left = .Left + 45
        .Top = .Top + 45
    End If
    .Height = flxZusatz.RowHeight(1)
    .Width = flxZusatz.ColWidth(0)  ' flxZusatz.Width - 90
    .Visible = True
    .ZOrder 0
    .SetFocus
End With

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseDown")
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
    
If (y <= wPara1.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseMove")
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
Dim c As Object

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is nlCommand) Then
        If (c.MouseOver) Then
            c.MouseOver = 0
        End If
    End If
Next
On Error GoTo DefErr

Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
If (iNewLine) And (Me.Visible) Then
    CurrentX = wPara1.NlFlexBackY
    CurrentY = (wPara1.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdF2_Click()
Call cmdF2_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    End If
End If

End Sub

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub


