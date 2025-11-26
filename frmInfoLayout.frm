VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInfoLayout 
   Caption         =   "Einstellung des Infobereichs"
   ClientHeight    =   4485
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   2
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   3
      Top             =   3600
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
      _Version        =   65541
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
      _Version        =   65541
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
End
Attribute VB_Name = "frmInfoLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "frmInfoLayout.frm"

Private Sub cmdEsc_Click()

Unload Me
End Sub

Private Sub cmdOk_Click()

'InfoLayoutInd% = Val(flxInfoLayout(1).TextMatrix(flxInfoLayout(1).row, 1))
InfoLayoutBelegung$ = flxInfoLayout(1).TextMatrix(flxInfoLayout(1).row, 1)
Unload Me
End Sub

Private Sub flxInfoLayout_DblClick(index As Integer)
cmdOk.Value = True
End Sub

Private Sub flxInfoLayout_RowColChange(index As Integer)

If (index = 0) Then
    If (flxInfoLayout(0).Redraw) Then
        Call InsertLayoutDateien
    End If
End If

End Sub

Private Sub Form_Load()
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim h$, h2$, FormStr$

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
    Else
        .AddItem Taxe1.DateiBezeichnung + vbTab + Taxe1.DateiKurz
        .AddItem Ast1.DateiBezeichnung + vbTab + Ast1.DateiKurz
        .AddItem Ass1.DateiBezeichnung + vbTab + Ass1.DateiKurz
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
    .row = 1
    For i% = 2 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = InfoLayoutDatei$) Then
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

'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

End Sub

Sub InsertLayoutDateien()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("InsertLayoutDateien")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, anz%, AktRow0%, row%
Dim h$
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
    If (row% = 1) Then
        For row% = 2 To (flxInfoLayout(0).Rows - 1)
            If (row% = 2) Then
                If (MatchTyp% = MATCH_LIEFERANTEN) Then
                    Set iClass = Lif1
                Else
                    Set iClass = Taxe1
                End If
            ElseIf (row% = 3) Then
                Set iClass = Ast1
            ElseIf (row% = 4) Then
                Set iClass = Ass1
            End If
            For j% = 0 To (iClass.AnzFelder - 1)
                h$ = clsDat.FirstLettersUcase$(RTrim(iClass.FeldBezeichnung(j%)))
                h$ = h$ + " (" + iClass.DateiBezeichnung + ")"
                h$ = h$ + vbTab + iClass.DateiKurz + "." + iClass.FeldKurz(j%)
                .AddItem h$
            Next j%
        Next row%
    Else
        If (row% = 2) Then
            If (MatchTyp% = MATCH_LIEFERANTEN) Then
                Set iClass = Lif1
            Else
                Set iClass = Taxe1
            End If
        ElseIf (row% = 3) Then
            Set iClass = Ast1
        ElseIf (row% = 4) Then
            Set iClass = Ass1
       End If
        For j% = 0 To (iClass.AnzFelder - 1)
            h$ = clsDat.FirstLettersUcase$(RTrim(iClass.FeldBezeichnung(j%)))
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
        
If (flxInfoLayout(0).TextMatrix(flxInfoLayout(0).row, 1) = InfoLayoutDatei$) Then
    With flxInfoLayout(1)
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 1) = InfoLayoutBelegung$) Then
                .row = i%
                .TopRow = i%
                .CellFontBold = True
                flxInfoLayout(0).CellFontBold = True
                Exit For
            End If
        Next i%
    End With
End If

'Call DefErrPop
End Sub

'Sub InsertLayoutDateien(DateiInd%)
'''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''Call DefErrFnc("InsertLayoutDateien")
''On Error GoTo DefErr
''GoTo DefErrEnd
''DefErr:
''Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
''Case vbRetry
''  Resume
''Case vbIgnore
''  Resume Next
''End Select
''Call DefErrAbort
''DefErrEnd:
'''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, anz%, AktRow0%
'Dim h$
'
'With flxInfoLayout(1)
'    .Redraw = False
'    .Rows = 2
'    .FixedRows = 1
'    If (DateiInd% = 0) Then
'        .FormatString = "<" + "Alle Dateien" + "|<"
'    Else
'        .FormatString = "<" + DateiTab(DateiInd% - 1).Bez + "|<"
'    End If
'    .Rows = 1
'    .Cols = 2
'    .ColWidth(0) = .Width
'    .ColWidth(1) = 0
'
'    .row = 0
'    .col = 0
'    .CellFontBold = True
'
'    If (DateiInd% = 0) Then
'        For i% = 0 To 5
'            anz% = DateiTab(i%).FeldAnz
'            For j% = 0 To (anz% - 1)
'                h$ = FirstLettersUcase$(RTrim(Feldtab(j%, Asc(DateiTab(i%).FTab)).Bez)) + "  (" + RTrim$(DateiTab(i%).Bez) + ")"
'                h$ = h$ + vbTab + Str$(i% * 100 + (j%))
'                .AddItem h$
'            Next j%
'        Next i%
'    Else
'        DateiInd% = DateiInd% - 1
'        anz% = DateiTab(DateiInd%).FeldAnz
'        For j% = 0 To (anz% - 1)
'            h$ = FirstLettersUcase$(RTrim(Feldtab(j%, Asc(DateiTab(DateiInd%).FTab)).Bez))
'            h$ = h$ + vbTab + Str$(DateiInd% * 100 + (j%))
'            .AddItem h$
'        Next j%
'    End If
'
'    .RowSel = .row
'    .col = 0
'    .ColSel = .Cols - 1
'    .Sort = flexSortStringNoCaseAscending
'    .row = 1
'    .TopRow = 1
'    .Redraw = True
'End With
'
'If (InfoLayoutInd% > 0) And (flxInfoLayout(0).row = (InfoLayoutInd% \ 100) + 2) Then
'    flxInfoLayout(0).CellFontBold = True
'    For i% = 1 To (flxInfoLayout(1).Rows - 1)
'        If (Val(flxInfoLayout(1).TextMatrix(i%, 1)) = InfoLayoutInd%) Then
'            flxInfoLayout(1).row = i%
'            flxInfoLayout(1).TopRow = i%
'            flxInfoLayout(1).CellFontBold = True
'            Exit For
'        End If
'    Next i%
'End If
'
''Call DefErrPop
'End Sub

Private Sub flxInfoLayout_KeyPress(index As Integer, KeyAscii As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("flxInfoLayout_KeyPress")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, row%, Gef%
Dim ch$

ch$ = UCase$(Chr$(KeyAscii))
If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    Gef% = False
    With flxInfoLayout(index)
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                .row = i%
                Gef% = True
                Exit For
            End If
        Next i%
        If (Gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                    .row = i%
                    Gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (Gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
        End If
    End With
End If

'Call DefErrPop
End Sub


