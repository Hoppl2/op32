VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmInfoLayout 
   Caption         =   "Einstellung des Infobereichs"
   ClientHeight    =   4095
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6480
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
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
      Left            =   5640
      Picture         =   "InfoLayout.frx":0000
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
      Left            =   5880
      Picture         =   "InfoLayout.frx":00A9
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
      Left            =   6120
      Picture         =   "InfoLayout.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxInfoLayout 
      Height          =   2700
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
   Begin MSFlexGridLib.MSFlexGrid flxInfoLayout 
      Height          =   2700
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   0
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
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmInfoLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "INFOLAYOUT.FRM"

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
Dim i%
Dim h$
        
With flxInfoLayout(1)
    .CellFontBold = True

    If (flxInfoLayout(0).TextMatrix(flxInfoLayout(0).row, 0) = clsKz.DateiBezeichnung) Then
        h$ = ""
        For i% = 1 To (.Rows - 1)
            .row = i%
            If (.CellFontBold) Then
                h$ = h$ + .TextMatrix(.row, 1) + " "
            End If
        Next i%
        InfoLayoutBelegung$ = h$
        InfoLayoutBezeichnung$ = "Kennzeichen"
    Else
        InfoLayoutBelegung$ = .TextMatrix(.row, 1)
        InfoLayoutBezeichnung$ = .TextMatrix(.row, 0)
    End If
End With

Unload Me

Call clsError.DefErrPop
End Sub

Private Sub flxInfoLayout_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfoLayout_DblClick")
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

Private Sub flxInfoLayout_RowColChange(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfoLayout_RowColChange")
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

If (Index = 0) Then
    If (flxInfoLayout(0).Redraw) Then
        Call InsertLayoutDateien
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
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim h$, h2$, FormStr$, iInfoLayoutDatei$

Call wPara1.InitFont(Me)

With flxInfoLayout(0)
    .Rows = 2
    .FixedRows = 1
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 11 + 90
    .Width = TextWidth("Wwwwwwwwwwwwwwwww")
    
    .FormatString = "<Datenquellen|"
    .ColWidth(0) = .Width
    .ColWidth(1) = 0
    .Rows = 1
    .AddItem "Alle Dateien"
    If (MatchTyp% = MATCH_LIEFERANTEN) Then
        .AddItem Lif1.DateiBezeichnung + vbTab + Lif1.DateiKurz
    ElseIf (MatchTyp% = MATCH_HILFSTAXE) Then
        .AddItem hTax1.DateiBezeichnung + vbTab + hTax1.DateiKurz
    Else
        If (TaxeAdoDBok) Then
            .AddItem TaxeAdoDB1.DateiBezeichnung + vbTab + TaxeAdoDB1.DateiKurz
        Else
            .AddItem Taxe1.DateiBezeichnung + vbTab + Taxe1.DateiKurz
        End If
        If (ArtikelDBok) Then
            .AddItem ArtikelDB1.DateiBezeichnung + vbTab + ArtikelDB1.DateiKurz
        End If
        .AddItem Ast1.DateiBezeichnung + vbTab + Ast1.DateiKurz
        .AddItem Ass1.DateiBezeichnung + vbTab + Ass1.DateiKurz
        .AddItem clsKz.DateiBezeichnung + vbTab + clsKz.DateiKurz
        .AddItem nnek1.DateiBezeichnung + vbTab + nnek1.DateiKurz
        If (iMARS) Then
            .AddItem "DocMorris-Daten" + vbTab + "DM"
        End If
'        .AddItem clsKz.DateiBezeichnung + vbTab + clsKz.DateiKurz
    End If
End With

With flxInfoLayout(1)
    .Rows = 2
    .FixedRows = 1
    
    .Top = wPara1.TitelY
    .Left = flxInfoLayout(0).Left + flxInfoLayout(0).Width + 300
    .Height = .RowHeight(0) * 11 + 90
    .Width = TextWidth("Wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww")
    
    .ColWidth(0) = .Width
    .Rows = 1
End With

With flxInfoLayout(0)
    iInfoLayoutDatei$ = InfoLayoutDatei$
'    If (ArtikelDBok) Then
'        If (iInfoLayoutDatei$ = "AST") Or (iInfoLayoutDatei$ = "ASS") Then
'            iInfoLayoutDatei = "ART"
'        End If
'    End If
    
    .row = 1
    For i% = 2 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = iInfoLayoutDatei$) Then
            .row = i%
            Exit For
        End If
    Next i%
End With
Call InsertLayoutDateien

Font.Bold = False   ' True

cmdOk.Top = flxInfoLayout(0).Top + flxInfoLayout(0).Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = flxInfoLayout(1).Left + flxInfoLayout(1).Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    For i = 0 To 1
        With flxInfoLayout(i)
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
    Next i
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdOk.Top
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    For i = 0 To 1
        flxInfoLayout(i).Top = flxInfoLayout(i).Top + iAdd2
    Next i
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdOk.Top
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = flxInfoLayout(0).Top + flxInfoLayout(0).Height + iAdd + 600
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

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450

    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxInfoLayout(0).Left - iAdd) / Screen.TwipsPerPixelX, (flxInfoLayout(0).Top - iAdd) / Screen.TwipsPerPixelY, (flxInfoLayout(1).Left + flxInfoLayout(1).Width + iAdd) / Screen.TwipsPerPixelX, (flxInfoLayout(0).Top + flxInfoLayout(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

'    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
'    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

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
    RoundRect hdc, (flxInfoLayout(0).Left - iAdd) / Screen.TwipsPerPixelX, (flxInfoLayout(0).Top - iAdd) / Screen.TwipsPerPixelY, (flxInfoLayout(1).Left + flxInfoLayout(1).Width + iAdd) / Screen.TwipsPerPixelX, (flxInfoLayout(0).Top + flxInfoLayout(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Sub InsertLayoutDateien()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InsertLayoutDateien")
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
Dim i%, j%, anz%, AktRow0%, row%, ind%
Dim h$, sKurz$
Dim iClass As Object

With flxInfoLayout(1)
    .Redraw = False
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<" + flxInfoLayout(0).text + "|<"
    .Rows = 1
    .Cols = 2
    .ColWidth(0) = .Width
    .ColWidth(1) = 0
    
    .row = 0
    .col = 0
    .CellFontBold = True
    
    row% = flxInfoLayout(0).row
    sKurz = flxInfoLayout(0).TextMatrix(row, 1)
    If (row% = 1) Then
        For row% = 2 To (flxInfoLayout(0).Rows - 1)
            sKurz = flxInfoLayout(0).TextMatrix(row, 1)
            If (row% = 2) Then
                If (MatchTyp% = MATCH_LIEFERANTEN) Then
                    If (LieferantenDBok%) Then
                        Set iClass = LieferantenDB1
                    Else
                        Set iClass = Lif1
                    End If
                ElseIf (MatchTyp% = MATCH_HILFSTAXE) Then
                    Set iClass = hTax1
                ElseIf (TaxeAdoDBok) Then
                    Set iClass = TaxeAdoDB1
                Else
                    Set iClass = Taxe1
                End If
            ElseIf (sKurz = "ART") Then
                If (ArtikelDBok%) Then
                    Set iClass = ArtikelDB1
                End If
            ElseIf (sKurz = "AST") Then
'                If (ArtikelDBok%) Then
'                    Set iClass = ArtikelDB1
'                Else
                    Set iClass = Ast1
'                End If
            ElseIf (sKurz = "ASS") Then
'                If (ArtikelDBok%) Then
'                    Set iClass = ArtikelDB1
'                Else
                    Set iClass = Ass1
'                End If
            ElseIf (sKurz = "KZ") Then
                Set iClass = clsKz
            ElseIf (sKurz = "NNEK") Then
'                If (ArtikelDBok%) Then
'                    Set iClass = ArtikelDB1
'                Else
                    Set iClass = nnek1
'                End If
'            ElseIf (row% = 3) Then
'                Set iClass = Ast1
'            ElseIf (row% = 4) Then
'                Set iClass = Ass1
'            ElseIf (row% = 6) Then
'                Set iClass = nnek1
            ElseIf (sKurz = "DM") Then
                Set iClass = DM1
'                End If
'            ElseIf (row% = 3) Then
'                Set iClass = Ast1
'            ElseIf (row% = 4) Then
'                Set iClass = Ass1
'            ElseIf (row% = 6) Then
'                Set iClass = nnek1
            End If
            For j% = 0 To (iClass.AnzFelder - 1)
'                h$ = clsDat.FirstLettersUcase$(RTrim(iClass.FeldBezeichnung(j%)))
                h$ = RTrim(iClass.FeldBezeichnung(j%))
                h$ = h$ + " (" + iClass.DateiBezeichnung + ")"
                h$ = h$ + vbTab + iClass.DateiKurz + "." + iClass.FeldKurz(j%)
                .AddItem h$
            Next j%
        Next row%
    Else
        If (row% = 2) Then
            If (MatchTyp% = MATCH_LIEFERANTEN) Then
                If (LieferantenDBok) Then
                    Set iClass = LieferantenDB1
                Else
                    Set iClass = Lif1
                End If
            ElseIf (MatchTyp% = MATCH_HILFSTAXE) Then
                Set iClass = hTax1
            ElseIf (TaxeAdoDBok) Then
                Set iClass = TaxeAdoDB1
            Else
                Set iClass = Taxe1
            End If
        ElseIf (sKurz = "ART") Then
            If (ArtikelDBok%) Then
                Set iClass = ArtikelDB1
            End If
        ElseIf (sKurz = "AST") Then
'            If (ArtikelDBok%) Then
'                Set iClass = ArtikelDB1
'            Else
                Set iClass = Ast1
'            End If
        ElseIf (sKurz = "ASS") Then
'            If (ArtikelDBok%) Then
'                Set iClass = ArtikelDB1
'            Else
                Set iClass = Ass1
'            End If
        ElseIf (sKurz = "KZ") Then
            Set iClass = clsKz
        ElseIf (sKurz = "NNEK") Then
            Set iClass = nnek1
            ElseIf (sKurz = "DM") Then
                Set iClass = DM1
'        ElseIf (row% = 3) Then
'            Set iClass = Ast1
'        ElseIf (row% = 4) Then
'            Set iClass = Ass1
'        ElseIf (row% = 5) Then
'            Set iClass = clsKz
'        ElseIf (row% = 6) Then
'            Set iClass = nnek1
        End If
        For j% = 0 To (iClass.AnzFelder - 1)
'            h$ = clsDat.FirstLettersUcase$(RTrim(iClass.FeldBezeichnung(j%)))
            h$ = RTrim(iClass.FeldBezeichnung(j%))
            h$ = h$ + vbTab + iClass.DateiKurz + "." + iClass.FeldKurz(j%)
            .AddItem h$
        Next j%
    End If
    
    .RowSel = .row
    .col = 0
    .ColSel = .Cols - 1
    .Sort = flexSortStringNoCaseAscending
    .row = 1
    .TopRow = 1
    .Redraw = True
End With
        
'If (ArtikelDBok) Then
'    If (InfoLayoutDatei$ = "AST") Or (InfoLayoutDatei$ = "ASS") Then
'        ind = InStr(InfoLayoutBelegung$, ".")
'        If (ind > 0) Then
'            sKurz = Mid(InfoLayoutBelegung, ind + 1)
'        End If
'        For j% = 0 To (ArtikelDB1.AnzFelderDos - 1)
''            h$ = clsDat.FirstLettersUcase$(RTrim(iClass.FeldBezeichnung(j%)))
'            h$ = RTrim(ArtikelDB1.FeldBezeichnungDos(j%))
'            If (h = sKurz) Then
'                InfoLayoutDatei = "ART"
'                InfoLayoutBelegung = ArtikelDB1.DateiKurz + "." + ArtikelDB1.FeldKurz(j% + 1)
'                Exit For
'            End If
'        Next j%
'    End If
'End If

If (flxInfoLayout(0).TextMatrix(flxInfoLayout(0).row, 1) = InfoLayoutDatei$) Then
    With flxInfoLayout(1)
        h$ = InfoLayoutBelegung$
        For i% = 1 To (.Rows - 1)
'            If (.TextMatrix(i%, 1) = InfoLayoutBelegung$) Then
            ind% = InStr(h$, .TextMatrix(i%, 1))
            If (ind% > 0) Then
                .row = i%
                .TopRow = i%
                .CellFontBold = True
                flxInfoLayout(0).CellFontBold = True
                
                h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1 + Len(.TextMatrix(i%, 1)))
                If (Trim$(h$) = "") Then Exit For
            End If
        Next i%
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub flxInfoLayout_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxInfoLayout_KeyPress")
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
Dim i%, row%, gef%
Dim ch$

ch$ = UCase$(Chr$(KeyAscii))

'If (iNewLine) Then
'    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
'        Call clsError.DefErrPop: Exit Sub
'    ElseIf (KeyAscii = 27) Then
'        Call nlcmdEsc_Click
'        Call clsError.DefErrPop: Exit Sub
'    End If
'End If

If (Index = 1) And (ch$ = " ") Then
    With flxInfoLayout(Index)
        If (.CellFontBold) Then
            .CellFontBold = False
        Else
            .CellFontBold = True
        End If
    End With
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxInfoLayout(Index)
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                    .row = i%
                    gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
        End If
    End With
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

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
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


