VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmWumsatzInfo 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4200
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3720
      Picture         =   "WumsatzInfo.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "WumsatzInfo.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3240
      Picture         =   "WumsatzInfo.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Default         =   -1  'True
      Height          =   450
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxWumsatzInfo 
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2540
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmWumsatzInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WUMSATZINFO.FRM"

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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, iToggle%, row%
Dim iAdd%, iAdd2%, x%, y%, wi%
Dim DRUCKHANDLE%
Dim zWertSumme#
Dim h$, h2$, sInfoTyp$, FormStr$

Call wPara1.InitFont(Me)

Call LifZus1.GetRecord(AktWumsatzLief% + 1)

With flxWumsatzInfo
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 1
    
    If (AktWumsatzInfo$ = "S") Then
        FormStr$ = "|>alle|>rabattfähig;|Sendungen/Zeitraum|Umsatz/Zeitraum|"
        FormStr$ = FormStr$ + "betrachtete Sendungen|betrachteter Umsatz|"
        FormStr$ = FormStr$ + "Umsatz/Sendung"
    Else
        FormStr$ = "|>alle|>rabattfähig;|erledigte Sendungen/Monat|geplante Sendungen/Monat|"
        FormStr$ = FormStr$ + "% Sendungen/Monat|Umsatz/Sendung|Umsatz bisher|"
        FormStr$ = FormStr$ + "Prognose Umsatz"
    End If
    .FormatString = FormStr$
    .SelectionMode = flexSelectionFree

    Font.Bold = True
    .ColWidth(0) = TextWidth("erledigte Sendungen/Monat")
    .ColWidth(1) = TextWidth("9 999 999.99")
    .ColWidth(2) = TextWidth("9 999 999.99")
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .Rows * .RowHeight(0) + 90
    
    row% = 1
    If (AktWumsatzInfo$ = "S") Then
        .TextMatrix(1, 1) = Format(LifZus1.OrgBeobachtSendungen, "0")
        .TextMatrix(3, 1) = Format(LifZus1.BeobachtSendungen, "0")
        For i% = 0 To 1
            .TextMatrix(2, 1 + i%) = Format(LifZus1.OrgBeobachtUmsatz(i%), "# ### ##0.00")
            .TextMatrix(4, 1 + i%) = Format(LifZus1.BeobachtUmsatz(i%), "# ### ##0.00")
            .TextMatrix(5, 1 + i%) = Format(LifZus1.UmsatzProSendung(i%), "# ### ##0.00")
        Next i%
    Else
        .TextMatrix(1, 1) = Format(LifZus1.SendungenAlt, "0")
        .TextMatrix(2, 1) = Format(LifZus1.SendungenPlan, "0")
        .TextMatrix(3, 1) = Format(CLng(LifZus1.AliquotProzent * 100#), "0") + "%"
        For i% = 0 To 1
            .TextMatrix(4, 1 + i%) = Format(LifZus1.UmsatzProSendung(i%), "# ### ##0.00")
            .TextMatrix(5, 1 + i%) = Format(LifZus1.UmsatzBisher(i%), "# ### ##0.00")
            .TextMatrix(6, 1 + i%) = Format(LifZus1.PrognoseUmsatz(i%), "# ### ##0.00")
        Next i%
    End If
    
    .FillStyle = flexFillRepeat
    .row = .Rows - 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .CellFontBold = True
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 1
    .HighLight = flexHighlightNever
    
'    Me.Caption = "Schwellwert-Automatik"
End With
    

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

Me.Width = flxWumsatzInfo.Left + flxWumsatzInfo.Width + 2 * wPara1.LinksX
Caption = AktWumsatzTyp$

With cmdEsc
    .Top = flxWumsatzInfo.Top + flxWumsatzInfo.Height + 150 * wPara1.BildFaktor
    .Width = wPara1.ButtonX%
    .Height = wPara1.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wPara1.TitelY% + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxWumsatzInfo
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
    
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxWumsatzInfo.Top = flxWumsatzInfo.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    
    Height = Height + iAdd2

    Me.Width = flxWumsatzInfo.Left + flxWumsatzInfo.Width + 600 * iFaktorX
    
    With nlcmdEsc
        .Init
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = flxWumsatzInfo.Top + flxWumsatzInfo.Height + 600 * iFaktorY
        .Top = .Top + iAdd
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

'    Me.Width = nlcmdEsc.Left + nlcmdEsc.Width + 600
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top)
'    RoundRect hdc, (flxWumsatzInfo.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatzInfo.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatzInfo.Left + flxWumsatzInfo.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatzInfo.Top + flxWumsatzInfo.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
Else
    nlcmdEsc.Visible = False
End If

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

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
    RoundRect hdc, (flxWumsatzInfo.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatzInfo.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatzInfo.Left + flxWumsatzInfo.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatzInfo.Top + flxWumsatzInfo.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

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

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdEsc_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    End If
End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub



