VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWuAuswahl 
   Caption         =   "Auswahl-Lieferungen"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6225
   Begin VB.ListBox lstSortierung 
      Height          =   255
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Neu (F2)"
      Height          =   540
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   2
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxWuAuswahl 
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
   Begin MSFlexGridLib.MSFlexGrid flxWuAuswahl 
      Height          =   2640
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmWuAuswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LieferantenRec
    Name As String * 12
    Nr As Byte
End Type

Private Type LieferungenRec
    LifDat As String * 11
    AnzArtikel As Integer
    Wert As Double
    fertig As Byte
    Name As String * 12
    Sort As String * 14
    RetourKz As String * 1
End Type

Dim AnzLieferungen%
Dim AlleLieferungen() As LieferungenRec

Dim AnzLieferanten%
Dim AlleLieferanten() As LieferantenRec


Private Const DefErrModul = "WUAUSWAHL.FRM"

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
Unload Me
Call DefErrPop
End Sub

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF2_Click")
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
Dim Lief%
Dim erg$, pzn$, txt$

erg$ = MatchCode(1, pzn$, txt$, False, False)
If (erg$ <> "") Then
    WuLifDat$ = "@" + Chr$(Val(pzn$)) + Format(Now, "DDMMYY")
    WuLifDat$ = WuLifDat$ + Format(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)), "0000") + " "
End If

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
Dim i%
Dim h$, h2$, LeftLif$, ErgLifDat$
        
If (WuAuswahlModus% = 2) Then
    If (ActiveControl.Name = flxWuAuswahl(0).Name) Then
        If (ActiveControl.Index = 0) Then
            flxWuAuswahl(1).SetFocus
            Call DefErrPop: Exit Sub
        End If
    End If
    With flxWuAuswahl(1)
        WuLifDat$ = .TextMatrix(.row, 7)
    End With
    Unload Me
    Call DefErrPop: Exit Sub
End If


If (WuAuswahlModus% = 0) Then
    ErgLifDat$ = WuLifDat$
Else
    ErgLifDat$ = RkLifDat$
End If

With flxWuAuswahl(0)
    If (.row = 1) Then
        LeftLif$ = Chr$(0)
    Else
        LeftLif$ = Chr$(AlleLieferanten(.row - 2).Nr)
    End If
End With

If (ActiveControl.Name = flxWuAuswahl(0).Name) Then
    If (ActiveControl.Index = 0) Then
        ErgLifDat$ = LeftLif$
    Else
        With flxWuAuswahl(1)
            ErgLifDat$ = ""
            If (.row = 1) Then
                ErgLifDat$ = LeftLif$ + " "
            ElseIf (.row = 2) And (WuAuswahlModus% = 0) Then
                ErgLifDat$ = LeftLif$ + "*"
            Else
                ErgLifDat$ = ErgLifDat$ + "@" + .TextMatrix(.row, 7) + .TextMatrix(.row, 0)
            End If
        End With
    End If
Else
    With flxWuAuswahl(1)
        h$ = ""
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 0) = Chr$(214)) Then
                h$ = h$ + "@" + .TextMatrix(i%, 7) + .TextMatrix(i%, 0)
            End If
        Next i%
    End With
    If (h$ <> "") Then ErgLifDat$ = h$
End If

If (WuAuswahlModus% = 0) Then
    WuLifDat$ = ErgLifDat$
Else
    RkLifDat$ = ErgLifDat$
End If

Unload Me

Call DefErrPop
End Sub

Private Sub flxwuauswahl_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuAuswahl_DblClick")
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

Private Sub flxWuAuswahl_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuAuswahl_GotFocus")
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
flxWuAuswahl(Index).HighLight = flexHighlightAlways
Call DefErrPop
End Sub

Private Sub flxWuAuswahl_lostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuAuswahl_lostFocus")
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
flxWuAuswahl(Index).HighLight = flexHighlightNever
Call DefErrPop
End Sub

Private Sub flxwuauswahl_RowColChange(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuAuswahl_RowColChange")
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
With flxWuAuswahl(Index)
    If (Index = 0) Then
        If (.redraw) Then
            Call InsertWuHeader
        End If
    End If
End With

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, Lief%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim h$, h2$, FormStr$

Call EinlesenWuHeader

Call wpara.InitFont(Me)

With flxWuAuswahl(0)
    .Rows = 2
    .FixedRows = 1
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 11 + 90
    .Width = TextWidth("Wwwwwwwwwwwwwwwww")
    
    .FormatString = "<Lieferanten|"
    .ColWidth(0) = .Width
    .ColWidth(1) = 0
    .Rows = 1
    .AddItem "Alle Lieferanten"
    For i% = 0 To (AnzLieferanten% - 1)
        .AddItem AlleLieferanten(i%).Name + vbTab + Str$(AlleLieferanten(i%).Nr)
    Next i%
    
End With

With flxWuAuswahl(1)
    .Rows = 2
    .FixedRows = 1
    .Rows = 1
    .FormatString = "|Lieferant|^Datum|^Uhrzeit|>Anz.Artikel|>Warenwert|||"
    
    Font.Bold = True
    .ColWidth(0) = TextWidth("X")
    .ColWidth(1) = TextWidth("XXXXXXXXXXXXXXX")
    .ColWidth(2) = TextWidth("99:99:9999")
    .ColWidth(3) = TextWidth("Uhrzeit ")
    .ColWidth(4) = TextWidth("Anz.Artikel ")
    .ColWidth(5) = TextWidth("Warenwert ")
    .ColWidth(6) = wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    .ColWidth(7) = 0
    .ColWidth(8) = 0
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Height = .RowHeight(0) * 11 + 90
    
    .Top = wpara.TitelY
    .Left = flxWuAuswahl(0).Left + flxWuAuswahl(0).Width + 300
End With

With flxWuAuswahl(0)
    .row = 1
End With
Call InsertWuHeader

Font.Bold = False   ' True

cmdOk.Top = flxWuAuswahl(0).Top + flxWuAuswahl(0).Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = flxWuAuswahl(1).Left + flxWuAuswahl(1).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With cmdF2
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = flxWuAuswahl(0).Left
    .Top = cmdOk.Top
    If (WuAuswahlModus% <> 0) Then .Visible = False
End With

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

h$ = "Auswahl - Lieferungen"
If (WuAuswahlModus% = 2) Then
    WuLifDat$ = ""
    h$ = h$ + "  für    " + KorrTxt$
End If
Me.Caption = h$

Call DefErrPop
End Sub

Sub InsertWuHeader()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InsertWuHeader")
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
Dim i%, j%, anz%, AktRow0%, row%, ind%, Lief%
Dim h$, sLief$, zeit$

With flxWuAuswahl(1)
    .redraw = False
    .Rows = 1
    .TextMatrix(0, 1) = flxWuAuswahl(0).text
    
    row% = flxWuAuswahl(0).row
    If (row% = 1) Then
        sLief$ = ""
    Else
        sLief$ = Chr$(AlleLieferanten(row% - 2).Nr)
    End If
    
    If (WuAuswahlModus% = 0) Then
        .AddItem vbTab + "alle akt. Lieferungen"
        .AddItem vbTab + "alle Altlasten"
    ElseIf (WuAuswahlModus% = 2) Then
    Else
        .AddItem vbTab + "alle akt. Rückkauf-Anfragen"
    End If
    
    For i% = 0 To (AnzLieferungen% - 1)
        If (sLief$ = "") Or (sLief$ = Left$(AlleLieferungen(i%).LifDat, 1)) Then
            .Rows = .Rows + 1
            row% = .Rows - 1
            
            h$ = " "
            If (AlleLieferungen(i%).fertig) Then
'                h$ = "?"
                h$ = AlleLieferungen(i%).RetourKz
            End If
            .TextMatrix(row%, 0) = h$
            .TextMatrix(row%, 1) = AlleLieferungen(i%).Name
            
            h$ = Mid$(AlleLieferungen(i%).LifDat, 2, 6)
            .TextMatrix(row%, 2) = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "." + Right$(h$, 2)
            
            zeit$ = Mid$(AlleLieferungen(i%).LifDat, 8, 4) ' Format(CVI(Right$(AlleLieferungen(i%).LifDat, 2)), "0000")
            .TextMatrix(row%, 3) = Left$(zeit$, 2) + ":" + Mid$(zeit$, 3)
            .TextMatrix(row%, 4) = Format(AlleLieferungen(i%).AnzArtikel, "0")
            .TextMatrix(row%, 5) = Format(AlleLieferungen(i%).Wert, "0.00")
            .TextMatrix(row%, 7) = AlleLieferungen(i%).LifDat
            .TextMatrix(row%, 8) = .TextMatrix(row%, 0)
        End If
    Next i%
    
    .TopRow = 1
    .row = 1
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
    .redraw = True
End With
        
Call DefErrPop
End Sub

Private Sub flxwuauswahl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxwuauswahl_KeyDown")
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

If (KeyCode = vbKeyF2) And (cmdF2.Visible) Then
    cmdF2.Value = True
End If

Call DefErrPop
End Sub

Private Sub flxwuauswahl_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuAuswahl_KeyPress")
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
Dim i%, row%, gef%, col%
Dim ch$, h$

ch$ = UCase$(Chr$(KeyAscii))

If (Index = 1) And (ch$ = " ") Then
    With flxWuAuswahl(Index)
        .col = 0
        If (.TextMatrix(.row, 0) = Chr$(214)) Then
'            .TextMatrix(.row, 0) = " "
            .TextMatrix(.row, 0) = .TextMatrix(.row, 8)
            h$ = ""
            For i% = 0 To .Cols - 1
                h$ = h$ + .TextMatrix(.row, i%) + vbTab
            Next i%
            row% = .row
            .redraw = False
            .RemoveItem row%
            .AddItem h$, row%
            .row = row%
            .redraw = True
        Else
            .CellFontName = "Symbol"
            .TextMatrix(.row, 0) = Chr$(214)
        End If
        .ColSel = .Cols - 1
    End With
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxWuAuswahl(Index)
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

Call DefErrPop
End Sub

Sub EinlesenWuHeader()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWuHeader")
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
Dim i%, j%, Max%, IstNeu%, Lief%, row%, ind%
Dim zWert#
Dim LifDat$, h$, h2$, zeit$, lName$, RetourKz$
Dim IstAltLast As Byte

AnzLieferungen% = 0
AnzLieferanten% = 0

If (WuAuswahlModus% = 0) Or (WuAuswahlModus% = 2) Then
    ww.GetRecord (1)
    Max% = ww.erstmax
Else
    rk.GetRecord (1)
    Max% = rk.erstmax
End If

For i% = 1 To Max%
    LifDat$ = ""
    RetourKz$ = "R"
    
    If (WuAuswahlModus% = 0) Or (WuAuswahlModus% = 2) Then
        ww.GetRecord (i% + 1)
    
        If (ww.status = 2) And (Asc(ww.WuBestDatum) <> 0) Then
            LifDat$ = Chr$(ww.Lief) + ww.WuBestDatum + Format(CVI(ww.WuBestZeit), "0000")
            zWert# = ww.WuAEP * ww.WuRm
            IstAltLast = ww.IstAltLast
            If (ww.WuNeuLm >= 0) Then RetourKz$ = "?"
        End If
    Else
        rk.GetRecord (i% + 1)
    
        If (rk.status = 2) And (Asc(rk.WuBestDatum) <> 0) Then
            LifDat$ = Chr$(rk.Lief) + rk.WuBestDatum + Format(CVI(rk.WuBestZeit), "0000")
            zWert# = rk.WuAEP * rk.WuRm
            IstAltLast = 0
        End If
    End If
        
    
    If (LifDat$ <> "") Then
        IstNeu% = True
        For j% = 0 To (AnzLieferungen% - 1)
            If (AlleLieferungen(j%).LifDat = LifDat$) Then
                IstNeu% = False
                Exit For
            End If
        Next j%
        
        If (IstNeu%) Then
            ReDim Preserve AlleLieferungen(AnzLieferungen%)
            AlleLieferungen(AnzLieferungen%).LifDat = LifDat$
            AlleLieferungen(AnzLieferungen%).AnzArtikel = 1
            AlleLieferungen(AnzLieferungen%).Wert = zWert#
            
            AlleLieferungen(AnzLieferungen%).fertig = IstAltLast
            
            AlleLieferungen(AnzLieferungen%).RetourKz = RetourKz$
            
            IstNeu% = True
            Lief% = Asc(Left$(LifDat$, 1))
            lif.GetRecord (Lief% + 1)
            h$ = Trim$(lif.kurz)
            If (h$ <> "") Then
                If (Asc(Left$(h$, 1)) < 32) Then
                    h$ = ""
                End If
            End If
            lName$ = h$ + " (" + Mid$(Str$(Lief%), 2) + ")"
            AlleLieferungen(AnzLieferungen%).Name = lName$
            
            
            h$ = Format(AlleLieferungen(AnzLieferungen%).fertig, "0")
            h2 = Mid$(LifDat$, 2, 6)
            h$ = h$ + Right$(h2$, 2) + Mid$(h2$, 3, 2) + Left$(h2$, 2)
            h2$ = Mid$(LifDat$, 8, 4)   'Format(CVI(Right$(LifDat$, 2)), "0000")
            h$ = h$ + h2$
            h2$ = Format(Lief%, "000")
            h$ = h$ + h2$
            AlleLieferungen(AnzLieferungen%).Sort = h$
            
            AnzLieferungen% = AnzLieferungen% + 1
            
            For j% = 0 To (AnzLieferanten% - 1)
                If (AlleLieferanten(j%).Nr = Lief%) Then
                    IstNeu% = False
                    Exit For
                End If
            Next j%
            
            If (IstNeu%) Then
                ReDim Preserve AlleLieferanten(AnzLieferanten%)
                AlleLieferanten(AnzLieferanten%).Nr = Lief%
                AlleLieferanten(AnzLieferanten%).Name = lName$
                AnzLieferanten% = AnzLieferanten% + 1
            End If
        Else
            AlleLieferungen(j%).AnzArtikel = AlleLieferungen(j%).AnzArtikel + 1
            AlleLieferungen(j%).Wert = AlleLieferungen(j%).Wert + zWert#
'            If (ww.IstAltLast) Then AlleLieferungen(j%).fertig = ww.IstAltLast
            If (ww.WuNeuLm >= 0) Then AlleLieferungen(j%).RetourKz = "?"
        End If
    End If
Next i%

If (AnzLieferanten% > 0) Then
    With lstSortierung
        .Clear
        For i% = 0 To (AnzLieferanten% - 1)
            .AddItem AlleLieferanten(i%).Name + vbTab + Str$(AlleLieferanten(i%).Nr)
        Next i%
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h$ = RTrim$(.text)
            ind% = InStr(h$, vbTab)
            AlleLieferanten(i%).Name = Left$(h$, ind% - 1)
            AlleLieferanten(i%).Nr = Val(Trim(Mid$(h$, ind% + 1)))
        Next i%
    
        
        .Clear
        For i% = 0 To (AnzLieferungen% - 1)
            h$ = AlleLieferungen(i%).Sort + AlleLieferungen(i%).RetourKz
'            h$ = h$ + vbTab + AlleLieferungen(i%).LifDat
            h$ = h$ + vbTab + Str$(AlleLieferungen(i%).AnzArtikel)
            h$ = h$ + vbTab + Format(AlleLieferungen(i%).Wert, "0.00")
            h$ = h$ + vbTab + AlleLieferungen(i%).Name
            .AddItem h$
        Next i%
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h$ = RTrim$(.text)
            
            AlleLieferungen(i%).fertig = Val(Left$(h$, 1))
            
            h2$ = Mid$(h$, 12, 3)
            LifDat$ = Chr$(Val(h2$))
            h2$ = Mid$(h$, 2, 6)
            LifDat$ = LifDat$ + Right$(h2$, 2) + Mid$(h2$, 3, 2) + Left$(h2$, 2)
            h2$ = Mid$(h$, 8, 4)
            LifDat$ = LifDat$ + h2$ 'mki(Val(h2$))
            AlleLieferungen(i%).LifDat = LifDat$    'Trim(Left$(h$, 9))
            AlleLieferungen(i%).RetourKz = Mid$(h$, 15, 1)

            h$ = Mid$(h$, 17)
            ind% = InStr(h$, vbTab)
            AlleLieferungen(i%).AnzArtikel = Val(Trim(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
            
            ind% = InStr(h$, vbTab)
            AlleLieferungen(i%).Wert = CDbl(Trim(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
            AlleLieferungen(i%).Name = Trim(h$)
        Next i%
    End With
End If

Call DefErrPop
End Sub


