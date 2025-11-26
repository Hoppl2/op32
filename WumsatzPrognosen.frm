VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmWumsatzPrognosen 
   Caption         =   "Prognose der Lieferantenumsätze"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4065
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3600
      Picture         =   "WumsatzPrognosen.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3360
      Picture         =   "WumsatzPrognosen.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3120
      Picture         =   "WumsatzPrognosen.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "&Details ... (F2)"
      Height          =   540
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxWumsatzPrognosen 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmWumsatzPrognosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WUMSATZPROGNOSEN.FRM"

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

frmWumsatz.Show 1

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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%, iAdd%, iAdd2%, x%, y%, wi%
Dim h$, h2$, FormStr$, SollStr$

Call wPara1.InitFont(Me)

With flxWumsatzPrognosen
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 1
    .Rows = 1
    
    FormStr$ = "|^Prognose Mindest-Umsatz|^Prognose Schwellwert"
    .FormatString = FormStr$
    .SelectionMode = flexSelectionFree
    
    Call PrognosenBefuellen
    If (.Rows = 1) Then .Rows = 2
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * .Rows + 90
    
    .ColWidth(0) = TextWidth("WWWWWW    ")
    .ColWidth(1) = TextWidth("Prognose Mindest-Umsatz  ")
    .ColWidth(2) = .ColWidth(1)
    .ColAlignment(0) = flexAlignLeftCenter

    spBreite% = 0
    For i% = 0 To (.Cols - 1)
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .row = 1
    .col = 1
End With


Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

With cmdF2
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
    .Left = flxWumsatzPrognosen.Left + flxWumsatzPrognosen.Width + 150 * wPara1.BildFaktor
    .Top = flxWumsatzPrognosen.Top
End With

Me.Width = cmdF2.Left + cmdF2.Width + 2 * wPara1.LinksX

With cmdEsc
    .Top = flxWumsatzPrognosen.Top + flxWumsatzPrognosen.Height + 150 * wPara1.BildFaktor
    .Width = wPara1.ButtonX%
    .Height = wPara1.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wPara1.TitelY% + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxWumsatzPrognosen
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
    
    cmdF2.Left = cmdF2.Left + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxWumsatzPrognosen.Top = flxWumsatzPrognosen.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    cmdF2.Top = cmdF2.Top + iAdd2
    
    Height = Height + iAdd2

    With nlcmdEsc
        .Init
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = flxWumsatzPrognosen.Top + flxWumsatzPrognosen.Height + 600 * iFaktorY
        .Top = .Top + iAdd
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
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top)
'    RoundRect hdc, (flxWumsatzPrognosen.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatzPrognosen.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatzPrognosen.Left + flxWumsatzPrognosen.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatzPrognosen.Top + flxWumsatzPrognosen.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
Else
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
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
    RoundRect hdc, (flxWumsatzPrognosen.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatzPrognosen.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatzPrognosen.Left + flxWumsatzPrognosen.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatzPrognosen.Top + flxWumsatzPrognosen.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub PrognosenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PrognosenBefuellen")
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
Dim i%, j%, SollRot%(1)
Dim proz&
Dim MindUms#, SchwellUms#, SchwellRab#
Dim h$, SQLStr$

With LifZus1
    For i% = 1 To .AnzRec
        .GetRecord (i% + 1)
        
        If (.WumsatzLief = 0) Or (.WumsatzLief = i%) Then
        
            MindUms# = .MindestUmsatz
            If (.TabTyp = 0) Then
                SchwellRab# = .Rabatt(0)
                SchwellUms# = .PrognoseSchwellwert
            Else
                SchwellRab# = 0#
                SchwellUms# = 0#
            End If
            
            If (MindUms# > 0#) Or (SchwellRab# > 0#) Then
                If (LieferantenDBok) Then
                    h = ""
                    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(i)
                    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                    If (LieferantenRec.RecordCount <> 0) Then
                        h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                    End If
                    LieferantenRec.Close
                ElseIf (i > 0) And (i <= Lif1.AnzRec) Then
                    Lif1.GetRecord (i% + 1)
                    h$ = Trim$(Lif1.kurz)
                End If
                If (h$ = String$(Len(h$), 0)) Then h$ = ""
                If (h$ = "") Then
                    h$ = "(" + Str$(i%) + ")"
                End If
                h$ = h$ + vbTab
                
                SollRot%(0) = False
                SollRot%(1) = False
                If (MindUms# > 0#) Then
                    proz& = CLng((.PrognoseUmsatz(0) / MindUms#) * 100#)
                    h$ = h$ + Format(proz&, "0") + "%"
                    If (proz& < 100) Then SollRot%(0) = True
                Else
                    h$ = h$ + " "
                End If
                h$ = h$ + vbTab
                
                If (SchwellRab# > 0#) Then
                    If (SchwellUms# = 0#) Then
                        proz& = 100
                    Else
                        proz& = CLng((.PrognoseUmsatz(1) / SchwellUms#) * 100#)
                    End If
                    h$ = h$ + Format(proz&, "0") + "%"
                    If (proz& < 100) Then SollRot%(1) = True
                Else
                    h$ = h$ + " "
                End If
                
                flxWumsatzPrognosen.AddItem h$
                
                For j% = 0 To 1
                    If (SollRot%(j%)) Then
                        flxWumsatzPrognosen.row = flxWumsatzPrognosen.Rows - 1
                        flxWumsatzPrognosen.col = j% + 1
                        flxWumsatzPrognosen.CellForeColor = vbRed
                    End If
                Next j%
            End If
        End If
    Next i%
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

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdF2_Click()
Call cmdF2_Click
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

If (iNewLine) Then
    If (KeyCode = vbKeyF2) Then
        nlcmdF2.Value = True
    End If
Else
    If (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
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



