VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmRezMonate 
   AutoRedraw      =   -1  'True
   Caption         =   "Historie Rezeptspeicher für "
   ClientHeight    =   5055
   ClientLeft      =   510
   ClientTop       =   375
   ClientWidth     =   4605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4605
   Begin VB.CommandButton cmdDruck 
      Caption         =   "Druck (F6)"
      Height          =   450
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   1200
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3000
      Picture         =   "RezMonate.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3240
      Picture         =   "RezMonate.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3480
      Picture         =   "RezMonate.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picFont 
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxRezSpeicher 
      Height          =   2700
      Left            =   360
      TabIndex        =   0
      Top             =   240
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
      Left            =   1920
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdDruck 
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmRezMonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Monate#(12, 3)

Dim iEditModus%


Private Const DefErrModul = "REZMONATE.FRM"

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
'Call ActProgram.RezHistorieExit

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
Dim h$

h$ = flxRezSpeicher.TextMatrix(flxRezSpeicher.row, 0)
If (h$ <> "") Then
    RezHistorieDatum$ = h$
    frmRezTage.Show 1
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

'Call AusdruckRezeptMonate

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
    End If
Else
    If KeyCode = vbKeyF6 Then
        Call cmdDruck_Click
        KeyCode = 0
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

Me.Left = frmRezSpeicher.Left
If (para.Newline) Then
    FormVersatzY% = wpara.NlCaptionY
    Me.Width = frmRezSpeicher.Width - 2 * wpara.NlFlexBackY
Else
    FormVersatzY% = wpara.FrmCaptionHeight + wpara.FrmBorderHeight
    Me.Width = frmRezSpeicher.Width
End If
Me.Height = frmRezSpeicher.Height - FormVersatzY%
Me.Top = frmRezSpeicher.Top + FormVersatzY%

'''''''''''''''''''''''''''''''''#
ScreenSizeWidth& = Me.ScaleWidth



With flxRezSpeicher
    .Rows = 2
    .FixedRows = 1
        
    .Cols = 13
    .FormatString = "|||||<Monat|>Anz.Rez.|>Ges.Wert|>Rab.Wert|>Zuzahlungen|>Abrechnung|>Herst.Rab.|>Wert/Rez.|>Artikel/Rez.|>"
    
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
        .ColWidth(6) = picFont.TextWidth(String(9, "9"))
        .ColWidth(7) = picFont.TextWidth(String(11, "9"))
        .ColWidth(8) = picFont.TextWidth(String(11, "9"))
        .ColWidth(9) = picFont.TextWidth(String(12, "9"))
        .ColWidth(10) = picFont.TextWidth(String(11, "9"))
        .ColWidth(11) = picFont.TextWidth(String(11, "9"))
        .ColWidth(12) = picFont.TextWidth(String(11, "9"))
        .ColWidth(13) = picFont.TextWidth(String(11, "9"))
        .ColWidth(14) = wpara.FrmScrollHeight
        
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

Caption = "Rezeptspeicher - " + RezHistorieKassenName$ + " 20" + RezHistorieDatum$

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
    cmdEsc.Visible = False

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

cmdDruck.Visible = False
nlcmdDruck.Visible = False

'Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

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
Dim h$, h2$, AktMo$, LastMo$
Dim i%, j%, Monat%
Dim dWert#, RabWerte#(12), HerstRabatt#(12)
Dim Jahr As Boolean
Dim bmk As Variant

For i% = 0 To 12
    For j% = 0 To 3
        Monate#(i%, j%) = 0#
    Next j%
Next i%

'MsgBox (RezHistorieDatum)
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
            
            If (Left(RezepteRec!DruckDatum, Len(RezHistorieDatum$)) = RezHistorieDatum$) Then
                If (RezHistorieIndexSuche%) Or (RezHistorieKassenNr$ = "") Or (Left$(RezepteRec!Kkasse, 7) <> "Privat ") Then
                    Monat% = Val(Mid$(RezepteRec!DruckDatum, 3, 2))
                    Monate#(Monat%, 0) = Monate#(Monat%, 0) + 1
                    Monate#(Monat%, 1) = Monate#(Monat%, 1) + RezepteRec!RezSumme
                    Monate#(Monat%, 2) = Monate#(Monat%, 2) + RezepteRec!RezGebSumme
                    Monate#(Monat%, 3) = Monate#(Monat%, 3) + RezepteRec!AnzArtikel
                    
                    If (IsNull(RezepteRec!RabattWert)) Then
                    Else
                        RabWerte#(Monat%) = RabWerte#(Monat%) + RezepteRec!RabattWert
                    End If
                    
                    If (IsNull(RezepteRec!HerstRabatt)) Then
                    Else
                        HerstRabatt#(Monat%) = HerstRabatt#(Monat%) + RezepteRec!HerstRabatt
                    End If
                End If
            End If
            
            .MoveNext
        Loop
        
        For i% = 1 To 12
            If (Monate#(i%, 0) > 0) Then
                For j% = 0 To 3
                    Monate#(0, j%) = Monate#(0, j%) + Monate#(i, j%)
                Next j%
                RabWerte#(0) = RabWerte#(0) + RabWerte#(i)
                HerstRabatt(0) = HerstRabatt(0) + HerstRabatt(i)
                
                h$ = Left(RezHistorieDatum$, 2) + Format(i%, "00")
                h$ = h$ + vbTab + vbTab + vbTab + vbTab + vbTab
                
                h2$ = Format(i%, "00")
                h$ = h$ + Format(CDate("01." + h2$ + ".20" + Left(RezHistorieDatum$, 2)), "MM/YYYY")
                
                
                h2$ = ""
                For j% = 0 To 2
                    h2$ = h2$ + vbTab
                    
                    If (Monate#(i%, j%) <> 0#) Then
                        If (j% = 0) Then
                            h2$ = h2$ + Format(Monate#(i%, j%), "0")
                        Else
                            h2$ = h2$ + Format(Monate#(i%, j%), "0.00")
                        End If
                    End If
                    
                    If (j% = 1) Then
                        h2$ = h2$ + vbTab
'                        dWert# = Monate#(i%, 1) / VmRabattFaktor#
                        dWert# = Monate#(i%, 1) - RabWerte#(i%)
                        h2$ = h2$ + Format(dWert#, "0.00")
                    ElseIf (j% = 2) Then
                        h2$ = h2$ + vbTab
                        dWert# = dWert# - Monate#(i%, 2)
                        If (dWert# <> 0#) Then
                            h2$ = h2$ + Format(dWert#, "0.00")
                        End If
                        
                        h2$ = h2$ + vbTab
                        dWert# = HerstRabatt#(i%)
                        If (dWert# <> 0#) Then
                            h2$ = h2$ + Format(dWert#, "0.00")
                        End If
                        
                        h2$ = h2$ + vbTab
                        If (Monate#(i%, 0) > 0) Then
                            dWert# = Monate#(i%, 1) / Monate#(i%, 0)
                            If (dWert# <> 0#) Then
                                h2$ = h2$ + Format(dWert#, "0.00")
                            End If
                        End If
                        
                        h2$ = h2$ + vbTab
                        If (Monate#(i%, 0) > 0) Then
                            dWert# = Monate#(i%, 3) / Monate#(i%, 0)
                            If (dWert# <> 0#) Then
                                h2$ = h2$ + Format(dWert#, "0.00")
                            End If
                        End If
                    End If
                Next j%
                        
                h$ = h$ + h2$
                flxRezSpeicher.AddItem h$
            End If
        Next i%
    
        i% = 0
        If (Monate#(i%, 0) > 0) Then
            h$ = ""
            h$ = h$ + vbTab + vbTab + vbTab + vbTab + vbTab
            
            h$ = h$ + ("20" + Left(RezHistorieDatum$, 2) + " ges.")
            
            
            h2$ = ""
            For j% = 0 To 2
                h2$ = h2$ + vbTab
                
                If (Monate#(i%, j%) <> 0#) Then
                    If (j% = 0) Then
                        h2$ = h2$ + Format(Monate#(i%, j%), "0")
                    Else
                        h2$ = h2$ + Format(Monate#(i%, j%), "0.00")
                    End If
                End If
                
                If (j% = 1) Then
                    h2$ = h2$ + vbTab
'                    dWert# = Monate#(i%, 1) / VmRabattFaktor#
                    dWert# = Monate#(i%, 1) - RabWerte#(i%)
                    h2$ = h2$ + Format(dWert#, "0.00")
                ElseIf (j% = 2) Then
                    h2$ = h2$ + vbTab
                    dWert# = dWert# - Monate#(i%, 2)
                    If (dWert# <> 0#) Then
                        h2$ = h2$ + Format(dWert#, "0.00")
                    End If
                    
                    h2$ = h2$ + vbTab
                    dWert# = HerstRabatt#(i%)
                    If (dWert# <> 0#) Then
                        h2$ = h2$ + Format(dWert#, "0.00")
                    End If
                        
                    h2$ = h2$ + vbTab
                    If (Monate#(i%, 0) > 0) Then
                        dWert# = Monate#(i%, 1) / Monate#(i%, 0)
                        If (dWert# <> 0#) Then
                            h2$ = h2$ + Format(dWert#, "0.00")
                        End If
                    End If
                    
                    h2$ = h2$ + vbTab
                    If (Monate#(i%, 0) > 0) Then
                        dWert# = Monate#(i%, 3) / Monate#(i%, 0)
                        If (dWert# <> 0#) Then
                            h2$ = h2$ + Format(dWert#, "0.00")
                        End If
                    End If
                End If
            Next j%
                    
            h$ = h$ + h2$
            flxRezSpeicher.AddItem h$
            
            With flxRezSpeicher
                .FillStyle = flexFillRepeat
                .row = .Rows - 1
                .col = 0
                .RowSel = .row
                .ColSel = .Cols - 1
                .CellBackColor = vbWhite
                .FillStyle = flexFillSingle
            End With
        End If
    End If
End With
    
With flxRezSpeicher
    If (.Rows = 1) Then .AddItem vbTab + vbTab + vbTab + vbTab + vbTab + "Keine Daten vorhanden!"
    .row = 1
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


