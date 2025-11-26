VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmBstatus 
   Caption         =   "Bestellstatus"
   ClientHeight    =   3675
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5640
   Icon            =   "Bstatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5640
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4680
      Picture         =   "Bstatus.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   4440
      Picture         =   "Bstatus.frx":0203
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   4200
      Picture         =   "Bstatus.frx":02B7
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxbstatus 
      Height          =   2280
      Left            =   600
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
      Left            =   2640
      TabIndex        =   2
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmBstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "BSTATUS.FRM"

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iAdd%, iAdd2%, x%, y%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)


Font.Bold = False   ' True

With flxbstatus
    If (AnzeigeFensterTyp$ = TEXT_BESTELLSTATUS) Then
        .Cols = 5
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Typ|^Datum|^Zeit|^Lieferant|^Menge"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("WWWWWWWWWWWWWWWW")
        .ColWidth(1) = TextWidth("99.99.9999")
        .ColWidth(2) = TextWidth("9999:99")
        .ColWidth(3) = TextWidth("WWWWWWW (999)")
        .ColWidth(4) = TextWidth("999999") + wPara1.FrmScrollHeight
    ElseIf (AnzeigeFensterTyp$ = TEXT_NACHBEARBEITUNG) Then
        .Cols = 10
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Datum|^Pzn|<Artikel|>Menge|<Eh|>BM|>NM|>RM|>LM|<User"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("99.99.9999")
        .ColWidth(1) = TextWidth(String(9, "9"))
        .ColWidth(2) = TextWidth(String(23, "W"))
        .ColWidth(3) = TextWidth(String(7, "9"))
        .ColWidth(4) = TextWidth(String(3, "W"))
        .ColWidth(5) = TextWidth("9999")
        .ColWidth(6) = TextWidth("9999")
        .ColWidth(7) = TextWidth("9999")
        .ColWidth(8) = TextWidth("9999")
        .ColWidth(9) = TextWidth(String(8, "W")) + wPara1.FrmScrollHeight
    ElseIf (AnzeigeFensterTyp$ = TEXT_ABSAGEN) Then
        .Cols = 6
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Datum|^Pzn|<Artikel|>Menge|<Eh|<BM"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("99.99.9999")
        .ColWidth(1) = TextWidth(String(9, "9"))
        .ColWidth(2) = TextWidth(String(23, "W"))
        .ColWidth(3) = TextWidth(String(7, "9"))
        .ColWidth(4) = TextWidth(String(3, "W"))
        .ColWidth(5) = TextWidth("9999") + wPara1.FrmScrollHeight
    End If

    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

Font.Bold = False   ' True

Me.Width = flxbstatus.Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
cmdOk.Top = flxbstatus.Top + flxbstatus.Height + 150

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxbstatus
'        .ScrollBars = flexScrollBarNone
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
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxbstatus.Top = flxbstatus.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = flxbstatus.Top + flxbstatus.Height + iAdd + 600 * iFaktorY
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .default = True
        .Cancel = True
        .Visible = True
    End With
    cmdOk.Visible = False

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxbstatus.Left - iAdd) / Screen.TwipsPerPixelX, (flxbstatus.Top - iAdd) / Screen.TwipsPerPixelY, (flxbstatus.Left + flxbstatus.Width + iAdd) / Screen.TwipsPerPixelX, (flxbstatus.Top + flxbstatus.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
Else
    nlcmdOk.Visible = False
End If

'Call BstatusBefuellen
'
'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

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
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
    RoundRect hdc, (flxbstatus.Left - iAdd) / Screen.TwipsPerPixelX, (flxbstatus.Top - iAdd) / Screen.TwipsPerPixelY, (flxbstatus.Left + flxbstatus.Width + iAdd) / Screen.TwipsPerPixelX, (flxbstatus.Top + flxbstatus.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

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

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdOk_Click
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

