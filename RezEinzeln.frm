VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmRezEinzeln 
   AutoRedraw      =   -1  'True
   Caption         =   "Historie Rezeptspeicher für "
   ClientHeight    =   6240
   ClientLeft      =   510
   ClientTop       =   375
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5295
   Begin VB.CommandButton cmdDruck 
      Caption         =   "Druck (F6)"
      Height          =   450
      Left            =   360
      TabIndex        =   14
      Top             =   5040
      Width           =   1200
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   2760
      Picture         =   "RezEinzeln.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3000
      Picture         =   "RezEinzeln.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3240
      Picture         =   "RezEinzeln.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picRezept 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox txtMarsRezeptAnmerkungen 
         Height          =   615
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "RezEinzeln.frx":0216
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox picFont 
      Height          =   495
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxRezSpeicher 
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin MSFlexGridLib.MSFlexGrid flxInfo 
      Height          =   540
      Index           =   0
      Left            =   3600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   953
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdDruck 
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblAv 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblAv 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblAv 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmRezEinzeln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "REZEINZELN.FRM"

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

If (picRezept.Visible) Then
    picRezept.Visible = False
Else
    Call ActProgram.RezHistorieExit
    Unload Me
End If

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
Dim Unique$

If (picRezept.Visible = False) Then
    Unique$ = flxRezSpeicher.TextMatrix(flxRezSpeicher.row, 0)
    RezepteRec.index = "Unique"
    RezepteRec.Seek "=", Unique$
    Call ActProgram.HoleAusRezeptSpeicher
    Call ActProgram.PaintRezept
End If

Call DefErrPop
End Sub

Private Sub cmdDruck_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdDruck_Click")
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

Call AusdruckRezeptEinzeln

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_GotFocus")
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

With flxRezSpeicher
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_LostFocus")
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

With flxRezSpeicher
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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
Dim OldRow&

If (para.Newline) Then
    If KeyCode = vbKeyF6 Then
        nlcmdDruck.Value = 1
        KeyCode = 0
'    ElseIf KeyCode = vbKeyF7 Then
'        nlcmdRezepte.Value = 1
'        KeyCode = 0
    End If
Else
    If KeyCode = vbKeyF6 Then
        Call cmdDruck_Click
        KeyCode = 0
'    ElseIf KeyCode = vbKeyF7 Then
'        Call cmdRezepte_Click
'        KeyCode = 0
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, FormVersatzY%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, OriginalBreite%
Dim iAdd%, iAdd2%
Dim ScreenSizeWidth&
Dim h$, h2$, h3$, FormStr$
Dim c As Control

iEditModus = 1

Call wpara.InitFont(Me)

Me.Left = frmRezTage.Left
If (para.Newline) Then
    FormVersatzY% = wpara.NlCaptionY
    Me.Width = frmRezTage.Width - 2 * wpara.NlFlexBackY
Else
    FormVersatzY% = wpara.FrmCaptionHeight + wpara.FrmBorderHeight
    Me.Width = frmRezTage.Width
End If
Me.Height = frmRezTage.Height - FormVersatzY%
Me.Top = frmRezTage.Top + FormVersatzY%
'''''''''''''''''''''''''''''''''#
ScreenSizeWidth& = Me.ScaleWidth



With flxRezSpeicher
    .Rows = 2
    .FixedRows = 1
        
    .Cols = 14
    .FormatString = "||||<Datum|<Rezept-Nr.|>Anz.Art.|>Ges.Wert|>Rab.Wert|>Zuzahlungen|>Abrechnung|>Herst.Rab.|>Zeit|<Kunde|<Bediener|>"
    
    .Rows = 1
    .Font.Size = .Font.Size + 1
    OriginalBreite = True
    Do
        If .Font.Size < 9 And .Font.Name <> "Small Fonts" Then
            .Font.Name = "Small Fonts"
            .Font.Size = 9
            picFont.FontName = .Font.Name
        End If
        If (.Font.Size - 1) <= 5 Then Exit Do
        .Font.Size = .Font.Size - 1
        picFont.FontSize = .Font.Size
        
        If OriginalBreite Then
          .ColWidth(5) = picFont.TextWidth(String(25, "A"))
          OriginalBreite = False
        Else
          .ColWidth(5) = picFont.TextWidth(String(15, "A"))
        End If
        
        For i% = 0 To 4
          .ColWidth(i%) = 0
        Next i%
        If (Len(RezHistorieDatum$) = 4) Then
            .ColWidth(4) = picFont.TextWidth(String(13, "9"))
        End If
        .ColWidth(6) = picFont.TextWidth(String(8, "9"))
        .ColWidth(7) = picFont.TextWidth(String(11, "9"))
        .ColWidth(8) = picFont.TextWidth(String(11, "9"))
        .ColWidth(9) = picFont.TextWidth(String(12, "9"))
        .ColWidth(10) = picFont.TextWidth(String(11, "9"))
        .ColWidth(11) = picFont.TextWidth(String(11, "9"))
        .ColWidth(12) = picFont.TextWidth(String(10, "9"))
        .ColWidth(13) = picFont.TextWidth(String(15, "9"))
        .ColWidth(14) = picFont.TextWidth(String(15, "9"))
        .ColWidth(15) = wpara.FrmScrollHeight
        
        Breite1% = 0
        For i% = 0 To (.Cols - 1)
            Breite1% = Breite1% + .ColWidth(i%)
        Next i%
        .Width = Breite1% + 90
    Loop While (.Width - 100) > ScreenSizeWidth&
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX

    .Height = ((Me.ScaleHeight - .Top - wpara.ButtonY - 300) \ .RowHeight(0)) * .RowHeight(0) + 90

    .Width = ScreenSizeWidth& - 2 * wpara.LinksX
    .ColWidth(5) = 0
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .ColWidth(5) = .Width - Breite1% - 90
    
    Call GetRezeptSpeicher
End With

With picRezept
    .Width = flxRezSpeicher.Width - flxRezSpeicher.ColWidth(5)
    .Height = .Width * 10.3 / 15
    If (.Height > flxRezSpeicher.Height) Then
        .Height = flxRezSpeicher.Height
        .Width = .Height * 15 / 10.3
    End If
    .Left = flxRezSpeicher.Left + flxRezSpeicher.ColPos(6)  ' wpara.LinksX
    If (Len(RezHistorieDatum) = 4) Then
        .Left = .Left - flxRezSpeicher.ColWidth(4)
    End If
    .Top = flxRezSpeicher.Top
    
    Call ActProgram.RezHistorieInit(Me)
'    Call ActProgram.PaintRezept
End With



Font.Bold = False   ' True


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdOk.Top = flxRezSpeicher.Top + flxRezSpeicher.Height + 150
cmdEsc.Top = cmdOk.Top

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

With cmdDruck
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption)
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = wpara.LinksX
End With

If (Len(RezHistorieDatum) = 4) Then
    h$ = Format(CDate("01." + Mid$(RezHistorieDatum$, 3, 2) + ".20" + Left(RezHistorieDatum$, 2)), "MM/YYYY")
    h$ = h$ + ", " + "Taxsumme über " + CStr(RezHistorieTaxSumme) + " EUR"
Else
    h$ = Format(CDate(Mid$(RezHistorieDatum$, 5, 2) + "." + Mid$(RezHistorieDatum$, 3, 2) + ".20" + Left(RezHistorieDatum$, 2)), "DD/MM/YYYY")
End If
Caption = "Rezeptspeicher - " + RezHistorieKassenName$ + " " + h$
                

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With nlcmdEsc
        .Init
    End With
    With flxRezSpeicher
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
        
        .Height = (Me.ScaleHeight - .Top - (iAdd + 600 + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450))
        .Height = (.Height \ .RowHeight(0)) * .RowHeight(0)
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    
    Height = Height + iAdd2
    
    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxRezSpeicher.Top + flxRezSpeicher.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdOk
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdEsc.Top
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdDruck
        .Init
        .AutoSize = True
        .Left = cmdDruck.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdDruck.Caption
        .TabIndex = cmdDruck.TabIndex
        .Enabled = cmdDruck.Enabled
        .Default = cmdDruck.Default
        .Cancel = cmdDruck.Cancel
        .Visible = True
    End With
    cmdDruck.Visible = False
    
    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxRezSpeicher
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (TypeOf c Is ComboBox) Then
                    Call wpara.ControlBorderless(c)
                ElseIf (c.Appearance = 1) Then
                    Call wpara.ControlBorderless(c, 2, 2)
                Else
                    Call wpara.ControlBorderless(c, 1, 1)
                End If

                If (c.Enabled) Then
                    c.BackColor = vbWhite
                Else
                    c.BackColor = Me.BackColor
                End If

'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor

                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
'            ElseIf (TypeOf c Is CheckBox) Then
'                c.Height = 0
'                c.Width = c.Height
'                If (c.Name = "chkHistorie") Then
'                    If (c.Index > 0) Then
'                        Load lblchkHistorie(c.Index)
'                    End If
'                    With lblchkHistorie(c.Index)
'                        .BackStyle = 0 'duchsichtig
'                        .Caption = c.Caption
'                        .Left = c.Left + 300
'                        .Top = c.Top
'                        .Width = TextWidth(.Caption) + 90
'                        .TabIndex = c.TabIndex
'                        .Visible = True
'                    End With
'                End If
            End If
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdDruck.Visible = False
End If
'''''''''

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_DblClick")
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
Call cmdOk_Click
Call DefErrPop
End Sub

'Sub RezHistorieBefuellen(ind%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("RezHistorieBefuellen")
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
'Dim i%, anz%
'Dim l&
'Dim Gesamt#, Fam#, ImpFähig#, ImpIst#
'Dim h$, SuchMonat$, SuchTag$
'
'If (ind% = 0) Then
'    flxRezHistorie(0).Rows = 1
'    With AuswertungRec
'        .Index = "Unique"
'        .Seek ">=", RezHistorieKassenNr$
'        If Not .NoMatch Then
'            Do While Not .EOF
'                If (AuswertungRec!Kkasse = RezHistorieKassenNr$) Then
'                    h$ = AuswertungRec!Monat
'                    h$ = h$ + vbTab + Format(CDate("01." + Mid(AuswertungRec!Monat, 3, 2) + ".20" + Left(AuswertungRec!Monat, 2)), "MM/YYYY")
'                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_Gesamt), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!RezAnzahl), "0")
'                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_GesamtFAM), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_ImpFähig), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_ImpIst), "0.00")
'                    h$ = h$ + vbTab
'                    flxRezHistorie(0).AddItem h$
'                Else
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'        Else
'            flxRezHistorie(0).AddItem " "
'        End If
'    End With
'    With flxRezHistorie(0)
'        .row = 1
'        .col = 0
'        .RowSel = .Rows - 1
'        .ColSel = .col
'        .Sort = 6
'        .col = 0
'        .ColSel = .Cols - 1
'    End With
'
'ElseIf (ind% = 1) Then
'    flxRezHistorie(1).Rows = 1
'    With RezepteRec
'        .Index = "Kasse"
'        SuchMonat$ = flxRezHistorie(0).TextMatrix(flxRezHistorie(0).row, 0)
'        .Seek ">=", RezHistorieKassenNr$, SuchMonat$ + "01"
'        If Not .NoMatch Then
'            Do While Not .EOF
'                If (RezepteRec!Kkasse = RezHistorieKassenNr$) And (Left$(RezepteRec!VerkDatum, 4) = SuchMonat$) Then
'                    h$ = vbTab + RezepteRec!VerkDatum
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!RezSumme), "0.00")
'                    h$ = h$ + vbTab + "1"
'    '                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!AnzArtikel), "0")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!Fam), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpFähig), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpIst), "0.00")
'                    h$ = h$ + vbTab
'                    flxRezHistorie(1).AddItem h$
'                Else
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'        Else
'            flxRezHistorie(1).AddItem " "
'        End If
'    End With
'    With flxRezHistorie(1)
'        .row = 1
'        .col = 1
'        .RowSel = .Rows - 1
'        .ColSel = .col
'        .Sort = 5
'
'        h$ = ""
'        l& = 1
'        Do
'            If (l& >= .Rows) Then Exit Do
'
'            If (.TextMatrix(l&, 1) = h$) Then
'                For i% = 2 To 6
'                    .TextMatrix(l& - 1, i%) = Format(xVal(.TextMatrix(l& - 1, i%)) + xVal(.TextMatrix(l&, i%)), "0.00")
'                Next i%
'
'                .RemoveItem l&
'            Else
'                h$ = .TextMatrix(l&, 1)
'                l& = l& + 1
'            End If
'        Loop
'    End With
'
'ElseIf (ind% = 2) Then
'    flxRezHistorie(2).Rows = 1
'    With RezepteRec
'        .Index = "Kasse"
'        SuchTag$ = flxRezHistorie(1).TextMatrix(flxRezHistorie(1).row, 1)
'        .Seek ">=", RezHistorieKassenNr$, SuchTag$
'        If Not .NoMatch Then
'            Do While Not .EOF
'                If (RezepteRec!Kkasse = RezHistorieKassenNr$) And (RezepteRec!VerkDatum = SuchTag$) Then
'                    h$ = RezepteRec!Unique + vbTab + RezepteRec!VerkDatum
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!RezSumme), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!AnzArtikel), "0")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!Fam), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpFähig), "0.00")
'                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpIst), "0.00")
'                    h$ = h$ + vbTab
'                    flxRezHistorie(2).AddItem h$
'                Else
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'        Else
'            flxRezHistorie(2).AddItem " "
'        End If
'    End With
'End If
'
'Call DefErrPop
'End Sub

Private Sub GetRezeptSpeicher()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetRezeptSpeicher")
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
Dim h$, h2$, AktMo$, LastMo$, sDatum$(1), sUhr$
Dim i%, j%, tag%, mon%, ok%
Dim lKuNr&
Dim dWert#, dRezSumme#
Dim Jahr As Boolean
Dim bmk As Variant
Dim abrRec As Recordset

Dim KundenDB As Database
Dim KundenRec As Recordset
Dim SQLStr$

Dim KundenSQL%
Dim KundenAdoRec As New ADODB.Recordset
Dim KundenAdoDB As New clsKundenDB

KundenSQL = KundenAdoDB.DBvorhanden
If (KundenSQL) Then
    KundenSQL = KundenAdoDB.OpenDB
Else
    Set KundenDB = OpenDatabase("Kunden.mdb", False, True)
End If

sDatum(0) = "01" + Mid(RezHistorieDatum$, 3, 2) + Mid(RezHistorieDatum$, 1, 2)
sDatum(1) = "31" + Mid(RezHistorieDatum$, 3, 2) + Mid(RezHistorieDatum$, 1, 2)
If (Len(RezHistorieDatum$) = 4) Then
    mon = Val(Mid$(RezHistorieDatum$, 3, 2))
    For i = 1 To 0 Step -1
        SQLStr$ = "SELECT * FROM AbrechnungsDaten WHERE Unique=" + CStr(mon)
        Set abrRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If Not (abrRec.EOF) Then
            sDatum(i) = abrRec!Datum
        End If
        
        mon = mon - 1
        If (mon < 1) Then
            mon = mon + 12
        End If
    Next i
End If
sDatum(0) = Mid$(sDatum(0), 5, 2) + Mid$(sDatum(0), 3, 2) + Mid$(sDatum(0), 1, 2)
sDatum(1) = Mid$(sDatum(1), 5, 2) + Mid$(sDatum(1), 3, 2) + Mid$(sDatum(1), 1, 2)
'MsgBox (RezHistorieDatum + "  " + sDatum(0) + "  " + sDatum(1))
    
With RezepteRec
    .index = "KasseDruck"
    If (RezHistorieIndexSuche%) Then
        .Seek ">=", RezHistorieKassenNr$, RezHistorieDatum$
    Else
        .Seek ">=", "", "02"
    End If
    If Not .NoMatch Then
        Do While Not .EOF
            If (RezHistorieIndexSuche%) Then
                If (RezepteRec!Kkasse <> RezHistorieKassenNr$) Or (Left(RezepteRec!DruckDatum, Len(RezHistorieDatum$)) <> RezHistorieDatum$) Then Exit Do
            End If
            
            If (Len(RezHistorieDatum$) = 4) Then
                h = RezepteRec!DruckDatum
                ok = (sDatum(0) < h) And (h <= sDatum(1))
                If (ok) Then
                    ok = (RezepteRec!RezSumme > RezHistorieTaxSumme)
                End If
            Else
                ok = (Left(RezepteRec!DruckDatum, Len(RezHistorieDatum$)) = RezHistorieDatum$)
            End If
                
            If (ok) Then
                
                
                If (RezHistorieIndexSuche%) Or (RezHistorieKassenNr$ = "") Or (Left$(RezepteRec!Kkasse, 7) <> "Privat ") Then
                    h2 = RezepteRec!DruckDatum
                    h2 = Mid$(h2, 5, 2) + "." + Mid$(h2, 3, 2) + ".20" + Mid$(h2, 1, 2)
                    
                    sUhr$ = Format(RezepteRec!DruckZeit \ 100, "00") + ":" + Format(RezepteRec!DruckZeit Mod 100, "00")
                    
                    h$ = RezepteRec!Unique + vbTab + RezepteRec!DruckDatum + sUhr + vbTab + vbTab + vbTab + h2$ + vbTab
                    
                    h2$ = RezepteRec!RezeptNr
                    If (Left$(h2$, 7) = "9999998") Then h2$ = Mid$(h2$, 8, 5)
                    
                    h$ = h$ + h2$
                    
                    h2$ = vbTab
                    h2$ = h2$ + Format(RezepteRec!AnzArtikel, "0")
                    h2$ = h2$ + vbTab + Format(RezepteRec!RezSumme, "0.00")
                   
                    h2$ = h2$ + vbTab
                    dRezSumme# = RezepteRec!RezSumme
                    If (IsNull(RezepteRec!RabattWert)) Then
                    Else
                        dRezSumme# = dRezSumme# - RezepteRec!RabattWert
                        h2$ = h2$ + Format(dRezSumme#, "0.00")
                    End If
                
'                    h2$ = h2$ + vbTab + Format(RezepteRec!RezSumme / VmRabattFaktor#, "0.00")
                    
                    h2$ = h2$ + vbTab + Format(RezepteRec!RezGebSumme, "0.00")
                    
'                    h2$ = h2$ + vbTab + Format(RezepteRec!RezSumme / VmRabattFaktor# - RezepteRec!RezGebSumme, "0.00")
                    h2$ = h2$ + vbTab + Format(dRezSumme# - RezepteRec!RezGebSumme, "0.00")

                    h2$ = h2$ + vbTab
                    If (IsNull(RezepteRec!HerstRabatt)) Then
                    Else
                        dWert# = RezepteRec!HerstRabatt
                        If (dWert# <> 0#) Then
                            h2$ = h2$ + Format(dWert#, "0.00")
                        End If
                    End If

'                    h2$ = h2$ + vbTab + Format(RezepteRec!VerkZeit / 100, "00") + ":" + Format(RezepteRec!VerkZeit Mod 100, "00")
                    h2$ = h2$ + vbTab + sUhr
                    h2$ = h2$ + vbTab
                    
                    On Error Resume Next
                    lKuNr = RezepteRec!knr
                    On Error GoTo DefErr
                    If (lKuNr > 0) Then
                        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(lKuNr)
                        If (KundenSQL) Then
                            On Error Resume Next
                            KundenAdoRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            KundenAdoRec.open SQLStr, KundenAdoDB.ActiveConn
                            If (KundenAdoRec.EOF) Then
                                h2$ = h2$ + "(" + CStr(lKuNr) + ")"
                            Else
                                h2$ = h2$ + Trim(CheckNullStr(KundenAdoRec!kurz))
                            End If
                        Else
                            Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
                            If (KundenRec.EOF) Then
                                h2$ = h2$ + "(" + CStr(lKuNr) + ")"
                            Else
                                h2$ = h2$ + Trim(CheckNullStr(KundenRec!kurz))
                            End If
                        End If
                    End If

                    h2$ = h2$ + vbTab
                    If (RezepteRec!PersonalNr% > 0) Then h2$ = h2$ + para.Personal(RezepteRec!PersonalNr)
                    
                    h$ = h$ + h2$
                    flxRezSpeicher.AddItem h$
                End If
            End If
            
            .MoveNext
        Loop
    Else
        flxRezSpeicher.AddItem " "
    End If
End With
        
If (KundenSQL) Then
    KundenAdoDB.CloseDB
Else
    KundenDB.Close
End If

With flxRezSpeicher
    If (.Rows = 1) Then .AddItem vbTab + vbTab + vbTab + vbTab + vbTab + "Keine Daten vorhanden!"
    
    .row = 1
    .col = 1    '11
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    .col = 0
    .ColSel = .Cols - 1
End With
    
Call DefErrPop
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseDown")
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
    
If (y <= wpara.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
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

Call DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Resize")
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

If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If

Call DefErrPop
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdDruck_click()
Call cmdDruck_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
'    ElseIf (KeyAscii = Asc("<")) And (nlcmdImport(0).Visible) Then
''        Call nlcmdChange_Click(0)
'        nlcmdImport(0).Value = 1
'    ElseIf (KeyAscii = Asc(">")) And (nlcmdImport(1).Visible) Then
''        Call nlcmdChange_Click(1)
'        nlcmdImport(1).Value = 1
    End If
End If
    
If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
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

Sub AusdruckRezeptEinzeln()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AusdruckRezeptEinzeln")
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
Dim ind%, ZeilenHöhe%, i%, j%, OrgAusrichtung%
Dim h$

AnzDruckSpalten% = flxRezSpeicher.Cols - 5
ReDim DruckSpalte(AnzDruckSpalten% - 1)

DruckSpalte(0).TypStr = String$(15, "9")
DruckSpalte(1).TypStr = String$(20, "9")
DruckSpalte(2).TypStr = String$(10, "9")
DruckSpalte(3).TypStr = String$(15, "9")
DruckSpalte(4).TypStr = String$(15, "9")
DruckSpalte(5).TypStr = String$(15, "9")
DruckSpalte(6).TypStr = String$(15, "9")
DruckSpalte(7).TypStr = String$(10, "9")
DruckSpalte(8).TypStr = String$(10, "9")
DruckSpalte(9).TypStr = String$(15, "9")
DruckSpalte(10).TypStr = String$(20, "9")

For i = 0 To (AnzDruckSpalten - 1)
    With DruckSpalte(i)
        .Ausrichtung = "L"
        
        h$ = flxRezSpeicher.TextMatrix(0, i + 4)
        ind = InStr("<^>", Left(h, 1))
        If (ind > 0) Then
            h = Mid(h, 2)
        End If
        .Titel = h$
        
'        If (ind = 2) Then
'            .Ausrichtung = "Z"
'        ElseIf (ind = 3) Then
'            .Ausrichtung = "R"
'        End If
        
        If (flxRezSpeicher.ColAlignment(i + 4) = flexAlignCenterCenter) Then
            .Ausrichtung = "Z"
        ElseIf (flxRezSpeicher.ColAlignment(i + 4) = flexAlignRightCenter) Then
            .Ausrichtung = "R"
        End If
    End With
Next i

OrgAusrichtung% = Printer.Orientation
Printer.Orientation = vbPRORLandscape
Call InitDruckZeile(True)

DruckSeite% = 0
Call RezEinzelnDruckKopf
ZeilenHöhe% = Printer.TextHeight("A")
'DruckSpalte(0).Attrib = 2
With flxRezSpeicher
    For i% = 1 To .Rows - 1
        h$ = .TextMatrix(i%, 4) + vbTab
        For j% = 5 To .Cols - 1
            If .ColWidth(j%) > 0 Then
                h$ = h$ + .TextMatrix(i%, j%) + vbTab
            End If
        Next j%
        Call DruckZeile(h$)
        If (Printer.CurrentY > Printer.ScaleHeight - 1000 - ZeilenHöhe%) Then
            Call DruckFuss
            Call RezEinzelnDruckKopf
        End If
    Next i%
End With
Call DruckFuss(False)
Printer.EndDoc
Printer.Orientation = OrgAusrichtung%

Call DefErrPop
End Sub

Sub RezEinzelnDruckKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezEinzelnDruckKopf")
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
Dim i%, x%, y%
Dim gesBreite&
Dim header$, KopfZeile$, Typ$, h$

'KopfZeile$ = Me.Caption
'header$ = KopfZeile$ '+ " " + cmbDatum.List(cmbDatum.ListIndex)
If (Len(RezHistorieDatum) = 4) Then
    h$ = Format(CDate("01." + Mid$(RezHistorieDatum$, 3, 2) + ".20" + Left(RezHistorieDatum$, 2)), "MM/YYYY")
Else
    h$ = Format(CDate(Mid$(RezHistorieDatum$, 5, 2) + "." + Mid$(RezHistorieDatum$, 3, 2) + ".20" + Left(RezHistorieDatum$, 2)), "DD/MM/YYYY")
End If
KopfZeile = "Rezeptspeicher - " + RezHistorieKassenName$ + " " + h$
header$ = KopfZeile$ '+ " " + cmbDatum.List(cmbDatum.ListIndex)
If (Len(RezHistorieDatum) = 4) Then
    header$ = header$ + ", " + "Taxsumme über " + CStr(RezHistorieTaxSumme) + " EUR"
End If
                
Call DruckKopf(header$, Typ$, KopfZeile$, 0)
Printer.CurrentY = Printer.CurrentY - 3 * Printer.TextHeight("A")
    
For i% = 0 To (AnzDruckSpalten% - 1)
    h$ = RTrim(DruckSpalte(i%).Titel)
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h$)
    End If
    Printer.CurrentX = x%
    Printer.Print h$;
Next i%

Printer.Print " "

y% = Printer.CurrentY
gesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, y%)-(gesBreite&, y%)

y% = Printer.CurrentY
Printer.CurrentY = y% + 30

Call DefErrPop

End Sub



