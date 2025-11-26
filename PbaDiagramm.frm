VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmPbaDiagramm 
   Caption         =   "Personal-Bedarfs-Analyse"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   Begin VB.CommandButton cmdF3 
      Caption         =   "&Planung (F3)"
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdF8 
      Caption         =   "&Vergleich (F8)"
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdF9 
      Caption         =   "&Speichern (F9)"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtEditAnalyse 
      BackColor       =   &H80000012&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "&Neu (F2)"
      Height          =   975
      Left            =   -120
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox cboAnalysen 
      Height          =   315
      Left            =   6600
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxBerechnungenGlobal 
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      _Version        =   393216
      Enabled         =   0   'False
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   0
   End
   Begin VB.PictureBox picAusdruck 
      Height          =   735
      Left            =   3480
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdF6 
      Caption         =   "&Drucken (F6)"
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid flxBerechnungen 
      Height          =   975
      Index           =   0
      Left            =   7200
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      _Version        =   393216
      Enabled         =   0   'False
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   0
   End
   Begin VB.CommandButton cmdF4 
      Caption         =   "F4"
      Height          =   975
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.ComboBox cmbWas 
      Height          =   315
      Index           =   0
      Left            =   8760
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox cmbTyp 
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown-Liste
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Esc"
      Height          =   975
      Left            =   10680
      TabIndex        =   8
      Top             =   6000
      Width           =   2175
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   2
      Left            =   -120
      OleObjectBlob   =   "PbaDiagramm.frx":0000
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   1
      Left            =   7560
      OleObjectBlob   =   "PbaDiagramm.frx":24CB
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   4575
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "PbaDiagramm.frx":4996
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   3
      Left            =   3600
      OleObjectBlob   =   "PbaDiagramm.frx":6E61
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   4
      Left            =   4920
      OleObjectBlob   =   "PbaDiagramm.frx":932C
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblVergleichAnalyse 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblAnalysen 
      Caption         =   "Gespeicherte &Analysen:"
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblBerechnungen 
      Caption         =   "Label1"
      Height          =   495
      Index           =   0
      Left            =   7080
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPbaDiagramm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SpInd%(50)
Dim AnzSp%

Dim StundenBez$(50)
Dim AnzStunden%

Dim AnzeigeModus%
Dim AnzeigeModusStr$(2)

Private Const DefErrModul = "PBADIAGRAMM.FRM"

Private Sub cmbTyp_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbTyp_Click")
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
Dim i%, ind%
Dim typ

Select Case cmbTyp(Index).ListIndex
Case 0
  typ = VtChChartType2dBar
Case 1
  typ = VtChChartType2dLine
Case 2
  typ = VtChChartType2dArea
Case 3
  typ = VtChChartType2dStep
Case 4
  typ = VtChChartType2dCombination
Case 5
  typ = VtChChartType2dPie
Case 6
  typ = VtChChartType2dXY
Case 7
  typ = VtChChartType3dBar
Case 8
  typ = VtChChartType3dLine
Case 9
  typ = VtChChartType3dArea
Case 10
  typ = VtChChartType3dStep
Case 11
  typ = VtChChartType3dCombination
End Select

chtPvs(Index).chartType = typ

PbaDiagrammTyp%(Index) = cmbTyp(Index).ListIndex
Call SpeicherIniPbaDiagramme(Index)
  
'With chtPvs(Index).DataGrid
'    .ColumnCount = AnzSp%
'    If (typ = VtChChartType2dBar) Then
'        ind% = cmbWas(Index).ListIndex
''        If (InStr(cmbWas(Index).Text, "/") > 0) Then
'        If (ind% >= 0) Then
'            If (dInfo$(ind%, 2, 0) = "#") Then
'                .ColumnCount = AnzSp% + 1
'                .ColumnLabel(AnzSp% + 1, 1) = "Gesamt"
'                .SetData 1, AnzSp% + 1, xVal(dInfo$(ind%, 1, 0)), False
'                chtPvs(Index).Plot.SeriesCollection(AnzSp% + 1).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
'            End If
'        End If
'    End If
'End With

Call DefErrPop
End Sub

Private Sub cmbWas_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbWas_Click")
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
Dim ind%, i%, j%, k%, col%
Dim iStundenPers%, iStundenPers2%, diff%
Dim KundenProStunde#, PersonalProStunde#, dStundenPers2#
  
With chtPvs(Index)
    ind% = cmbWas(Index).ListIndex
    With .DataGrid
    
        If (ind% = 0) Or (ind% = 2) Then
            .RowCount = 6
            .ColumnCount = AnzStunden%
            
            For k% = 1 To .RowCount
                If (Index = 0) Then
                    .RowLabel(k%, 1) = TagName$(k%)
                Else
                    .RowLabel(k%, 1) = Left$(TagName$(k%), 2)
                End If
            Next k%
            
            For k% = 1 To .ColumnCount
                .ColumnLabel(k%, 1) = StundenBez$(k% - 1)
            Next k%
        Else
            .RowCount = AnzStunden%
            .ColumnCount = 6
            
            For k% = 1 To .RowCount
                If (Index = 0) Then
                    .RowLabel(k%, 1) = StundenBez$(k% - 1)
                Else
                    .RowLabel(k%, 1) = Left$(StundenBez$(k% - 1), 2)
                End If
            Next k%
            
            For k% = 1 To .ColumnCount
                .ColumnLabel(k%, 1) = TagName$(k%)
            Next k%
        End If
         
         For j% = 1 To 6
            col% = 1
            For i% = 0 To 23
                KundenProStunde# = PbaRec(j%).KundenProStunde(i%)
                If (KundenProStunde# > 0) Then
                    PersonalProStunde# = PbaRec(j%).PersonalProStunde#(i%)
                    dStundenPers2# = PbaRec2(j%).PersonalProStunde#(i%)
                    
                    iStundenPers% = PbaRec(j%).iPersonalProStunde(i%)
                    iStundenPers2% = PbaRec2(j%).iPersonalProStunde(i%)
                Else
                    PersonalProStunde# = 0
                    dStundenPers2# = 0
                    
                    iStundenPers% = 0
                    iStundenPers2% = 0
                End If
                    
                    If (ind% = 0) And (col% <= .ColumnCount) Then
                        .SetData j%, col%, iStundenPers%, False
                    ElseIf (ind% = 1) And (col% <= .RowCount) Then
                        .SetData col%, j%, iStundenPers%, False
                    ElseIf (ind% = 2) And (col% <= .ColumnCount) Then
                        .SetData j%, col%, KundenProStunde#, False
                    ElseIf (col% <= .RowCount) Then
                        .SetData col%, j%, KundenProStunde#, False
                    End If
'                    LOCATE row%, 13 + (j% - 1) * 12
'                    If (AnzeigeTyp% = 0) Then
'                        If (Vergleich%) Then
'                            Diff% = CInt((dStundenPers2# / PersonalProStunde#) * 100) - 100
'                            If (dStundenPers2# = 0#) Then
'                              Print "  kA  ";
'                            ElseIf (Abs(Diff%) <= Prozent%) Then
'                              Print "      ";
'                            Else
'                              Print SetUsing$("####% ", CDbl(Diff%));
'                            End If
'                        Else
'                            Print SetUsing$("###.##", PersonalProStunde#);
'                        End If
'                    ElseIf (AnzeigeTyp% = 1) Then
'                        iStundenPers% = PbaRec(j%).iPersonalProStunde(i%)
'                        iStundenPers2% = PbaRec2(j%).iPersonalProStunde(i%)
'
'                        If (Vergleich%) Then
'                            Diff% = iStundenPers2% - iStundenPers%
'                            If (Diff% = 0) Then
'                                Print "      ";
'                            Else
'                                Print SetUsing$("  ### ", CDbl(Diff%));
'                            End If
'                        Else
'                            .SetData j%, col%, iStundenPers%, False
'                        End If
'                    ElseIf (AnzeigeTyp% = 2) Then
'                        If (Vergleich%) Then
'                            Diff% = CInt((PbaRec2(j%).KundenProStunde(i%) / KundenProStunde#) * 100) - 100
'                            If (PbaRec2(j%).KundenProStunde(i%) = 0#) Then
'                                Print "  kA  ";
'                            ElseIf (Abs(Diff%) <= Prozent%) Then
'                                Print "      ";
'                            Else
'                                Print SetUsing$(" ###% ", CDbl(Diff%));
'                            End If
'                            '          PRINT SetUsing$(" ####", PbaRec2(j%).KundenProStunde(i%) / KundenProStunde#);
'                        Else
'                            Print SetUsing$(" ####", KundenProStunde#);
'                        End If
'                    ElseIf (AnzeigeTyp% = 3) Then
'                        AnzHalbe% = CInt((KundenProStunde# / MaxSchnitt#) * 16)
'                        Print String$(AnzHalbe% \ 2, 219);
'                        If (AnzHalbe% Mod 2) Then
'                            Print Chr$(221);
'                        End If
'                    End If
'                End If
                If (i% >= gVon%) And (i% <= gBis%) Then
                    col% = col% + 1
                End If
            Next i%
        Next j%
             
'        For i% = 1 To .ColumnCount
'            .SetData 1, i%, CInt(xVal(dInfo$(ind%, 1, SpInd%(i% - 1)))), False
'        Next i%
'
'        If (cmbTyp(Index).ListIndex = 0) Then
''            If (InStr(cmbWas(Index).Text, "/") > 0) Then
'            If (dInfo$(ind%, 2, 0) = "#") Then
'                .ColumnCount = AnzSp% + 1
'                .ColumnLabel(AnzSp% + 1, 1) = "Gesamt"
'                .SetData 1, AnzSp% + 1, xVal(dInfo$(ind%, 1, 0)), False
'                chtPvs(Index).Plot.SeriesCollection(AnzSp% + 1).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
'            End If
'        End If

    End With
    PbaDiagrammWas%(Index) = ind%
    Call SpeicherIniPbaDiagramme(Index)
End With

Call DefErrPop
End Sub

Private Sub cmdF4_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF4_Click")
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

If (cmdF4.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

AnzeigeModus% = (AnzeigeModus% + 1) Mod 3

cmdF4.Caption = AnzeigeModusStr$(AnzeigeModus%)

If (AnzeigeModus% = 0) Then
    For i% = 0 To 4
        chtPvs(i%).Visible = False
        cmbTyp(i%).Visible = False
        cmbWas(i%).Visible = False
    Next i%
    For i% = 0 To 2
        flxBerechnungen(i%).Visible = True
        lblBerechnungen(i%).Visible = True
    Next i%
    cboAnalysen.Visible = True
    flxBerechnungenGlobal.Visible = True
    
    cmdF2.Enabled = True
    cmdF3.Enabled = True
    cmdF6.Enabled = True
    cmdF8.Enabled = True
    cmdF9.Enabled = True

ElseIf (AnzeigeModus% = 1) Then
    For i% = 1 To 4
        chtPvs(i%).Visible = False
        cmbTyp(i%).Visible = False
        cmbWas(i%).Visible = False
    Next i%
    For i% = 0 To 2
        flxBerechnungen(i%).Visible = False
        lblBerechnungen(i%).Visible = False
    Next i%
    cboAnalysen.Visible = False
    flxBerechnungenGlobal.Visible = False
    
    cmdF2.Enabled = False
    cmdF3.Enabled = False
    cmdF6.Enabled = False
    cmdF8.Enabled = False
    cmdF9.Enabled = False

    chtPvs(0).Visible = True
    cmbTyp(0).Visible = True
    cmbWas(0).Visible = True
    cmbTyp(0).SetFocus
    
Else
    For i% = 1 To 4
        chtPvs(i%).Visible = True
        cmbTyp(i%).Visible = True
        cmbWas(i%).Visible = True
    Next i%
    
    cmdF2.Enabled = False
    cmdF3.Enabled = False
    cmdF6.Enabled = False
    cmdF8.Enabled = False
    cmdF9.Enabled = False
    
    chtPvs(0).Visible = False
    cmbTyp(0).Visible = False
    cmbWas(0).Visible = False
    cmbTyp(1).SetFocus
End If

Call DefErrPop
End Sub

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbEsc_Click")
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

If (txtEditAnalyse(0).Visible) Then
    For i% = 1 To 2
        Unload txtEditAnalyse(i%)
    Next i%
    txtEditAnalyse(0).Visible = False
    
    PbaTest% = False

    cmdF2.Enabled = True
    cmdF3.Enabled = True
    cmdF4.Enabled = True
    cmdF6.Enabled = True
    cmdF8.Enabled = True
    cmdF9.Enabled = True
    cboAnalysen.Enabled = True
    cboAnalysen.SetFocus
Else
    Unload Me
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
Dim i%, j%, k%, col%, iStundenPers%, iStundenPers2%, gef%
Dim LegendenPos&
Dim a$, b2$, ab$, Bis$, h$
Dim KundenProStunde#, PersonalProStunde#, dStundenPers2#, SummeTag#

Call wpara.InitFont(Me)

AnzeigeModus% = 0

AnzeigeModusStr$(0) = "Graphik groß (F4)"
AnzeigeModusStr$(1) = "Graphik Übersicht (F4)"
AnzeigeModusStr$(2) = "Berechnungen (F4)"

Me.Width = frmAction.Width - 2 * wpara.FrmCaptionHeight
Me.Height = frmAction.Height - 2 * wpara.FrmCaptionHeight
Me.Left = frmAction.Left + wpara.FrmCaptionHeight
Me.Top = frmAction.Top + wpara.FrmCaptionHeight

For i% = 1 To 4
    Load cmbTyp(i%)
    Load cmbWas(i%)
Next i%
For i% = 1 To 2
    Load flxBerechnungen(i%)
    Load lblBerechnungen(i%)
Next i%

AnzStunden% = 0
For i% = 0 To 23
    If (i% >= gVon%) And (i% <= gBis%) Then
        If (i% = gVon%) Then
            ab$ = Format(gAbOffen%, "0000")
            Bis$ = Format(i% * 100 + 59, "0000")
        ElseIf (i% = gBis%) Then
            ab$ = Format(i% * 100, "0000")
            Bis$ = Format(gBisOffen%, "0000")
        Else
            a$ = Format(i%, "00")
            ab$ = a$ + "00"
            Bis$ = a$ + "59"
        End If
        b2$ = ab$ + "-" + Bis$
        StundenBez$(AnzStunden%) = b2$
        AnzStunden% = AnzStunden% + 1
    End If
Next i%

Call InitChartPositions

LegendenPos& = VtChLocationTypeLeft
If (LegendenPosStr$ = "R") Then
    LegendenPos& = VtChLocationTypeRight
ElseIf (LegendenPosStr$ = "O") Then
    LegendenPos& = VtChLocationTypeTop
ElseIf (LegendenPosStr$ = "U") Then
    LegendenPos& = VtChLocationTypeBottom
ElseIf (LegendenPosStr$ = "A") Then
    LegendenPos& = -1
End If

For j% = 0 To 4
    With chtPvs(j%)
    
        With .Legend.Location
            If (LegendenPos& = -1) Then
                .Visible = False
            Else
                .Visible = True
                .LocationType = LegendenPos& ' VtChLocationTypeLeft
            End If
        End With
        
        With .DataGrid
            .ColumnCount = AnzStunden%
            
            .RowCount = 6
            For i% = 1 To .RowCount
                .RowLabel(i%, 1) = TagName$(i%)
            Next i%
            For i% = 1 To .ColumnCount
                .ColumnLabel(i%, 1) = StundenBez$(i% - 1)
            Next i%
            
            
'            For i% = 1 To chtPvs(j%).Plot.SeriesCollection.Count
'                With chtPvs(j%).Plot.SeriesCollection(i%).DataPoints(-1)
'                    h$ = PersonalFarben$(SpInd%(i% - 1))
'                    b% = Val("&H" + Left$(h$, 2))
'                    g% = Val("&H" + Mid$(h$, 3, 2))
'                    r% = Val("&H" + Mid$(h$, 5, 2))
'                    .Brush.FillColor.Set r%, g%, b%
''                    .Brush.FillColor.Set Val(d"&H" + Left$(PersonalFarben$(SpInd%(i% - 1)), 2)), Val("&H" + Mid$(PersonalFarben$(SpInd%(i% - 1)), 3, 2)), Val("&H" + Mid$(PersonalFarben$(SpInd%(i% - 1)), 5, 2))
'                End With
''                chtPvs(j%).Plot.SeriesCollection(i%).DataPoints(0).Marker
'            Next
        End With
    End With
Next j%

For i% = 0 To 4
    With cmbTyp(i%)
        .AddItem "Balken 2D"
        .AddItem "Linie 2D"
        .AddItem "Fläche 2D"
        .AddItem "Schnitt 2D"
        .AddItem "Kombination 2D"
        .AddItem "Kreis 2D"
        .AddItem "X Y (Punkt) 2D"
        
        .AddItem "Säule 3D"
        .AddItem "Linie 3D"
        .AddItem "Fläche 3D"
        .AddItem "Schritt 3D"
        .AddItem "Kombination 3D"
        
        .ListIndex = PbaDiagrammTyp%(i%)
    End With
    
    With cmbWas(i%)
        .AddItem "Besetzung (Tage)"
        .AddItem "Besetzung (Stunden)"
        .AddItem "Kunden (Tage)"
        .AddItem "Kunden (Stunden)"
'        .ListIndex = 0
        .ListIndex = PbaDiagrammWas%(i%)
    End With
    
Next i%


For k% = 0 To 2
    With flxBerechnungen(k%)
         For j% = 1 To 6
            col% = 1
            For i% = 0 To 23
                KundenProStunde# = PbaRec(j%).KundenProStunde(i%)
                If (KundenProStunde# > 0) Then
                    PersonalProStunde# = PbaRec(j%).PersonalProStunde#(i%)
                    dStundenPers2# = PbaRec2(j%).PersonalProStunde#(i%)
                    
                    iStundenPers% = PbaRec(j%).iPersonalProStunde(i%)
                    iStundenPers2% = PbaRec2(j%).iPersonalProStunde(i%)
                Else
                    PersonalProStunde# = 0
                    dStundenPers2# = 0
                    
                    iStundenPers% = 0
                    iStundenPers2% = 0
                End If
                
                If (KundenProStunde# > 0#) Then
                    If (k% = 0) Then
                        h$ = Format(PersonalProStunde#, "0.0")
                    ElseIf (k% = 1) Then
                        h$ = Format(iStundenPers%, "0")
                    Else
                        h$ = Format(KundenProStunde#, "0")
                    End If
                    If (col% <= AnzStunden%) Then
                        .TextMatrix(col%, j%) = h$
                    End If
                End If
                    
                If (i% >= gVon%) And (i% <= gBis%) Then
                    col% = col% + 1
                End If
            Next i%
        
            SummeTag# = PbaRec(j%).KundenProTag
            If (SummeTag# > 0) Then
                If (k% = 0) Then
                    h$ = Format(PbaRec(j%).PersonalProTag, "0.0")
                ElseIf (k% = 1) Then
                    h$ = Format(PbaRec(j%).iPersonalProTag, "0")
                Else
                    h$ = Format(SummeTag#, "0")
                End If
                .TextMatrix(.Rows - 1, j%) = h$
            End If
        Next j%
        
        .FillStyle = flexFillRepeat
        .row = .Rows - 1
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellFontBold = True
        .FillStyle = flexFillSingle
    End With
Next k%
             
With flxBerechnungenGlobal
    ab$ = sDate(StartDatum%)
    h$ = Left$(ab$, 2) + "." + Mid$(ab$, 3, 2) + "." + Mid$(ab$, 5, 2)
    Bis$ = sDate(StopDatum%)
    h$ = h$ + " - " + Left$(Bis$, 2) + "." + Mid$(Bis$, 3, 2) + "." + Mid$(Bis$, 5, 2)
    .TextMatrix(0, 0) = "Beobachtungs-Zeitraum"
    .TextMatrix(1, 0) = "Personalstunden /  Woche"
    .TextMatrix(2, 0) = "Kundenanzahl / Tag"
'    .TextMatrix(3, 0) = "Leistungspotential des Personals"
    .TextMatrix(3, 0) = "Kunden / Normtag"
    .TextMatrix(0, 1) = h$
    .TextMatrix(1, 1) = Format(PersonalWochenStunden%, "0")
    .TextMatrix(2, 1) = Format(KundenProTag#, "0")
'    .TextMatrix(3, 1) = Format(KundenProArbeitsStunde#, "0")
    .TextMatrix(3, 1) = Format(KundenProArbeitsStunde# * 8, "0")
End With

With cmdF2
    .Width = TextWidth(.Caption) + 300
    .Height = wpara.ButtonY
    .Left = chtPvs(0).Left
    .Top = chtPvs(3).Top + chtPvs(3).Height + 90
End With
With cmdF3
    .Width = TextWidth(.Caption) + 300
    .Height = wpara.ButtonY
    .Left = cmdF2.Left + cmdF2.Width + 90
    .Top = cmdF2.Top
End With
With cmdF4
    .Width = TextWidth(AnzeigeModusStr$(1)) + 300
    .Caption = AnzeigeModusStr$(AnzeigeModus%)
    .Height = wpara.ButtonY
    .Left = cmdF3.Left + cmdF3.Width + 90
    .Top = cmdF3.Top
End With
With cmdF6
    .Width = TextWidth(.Caption) + 300
    .Height = wpara.ButtonY
    .Left = cmdF4.Left + cmdF4.Width + 90
    .Top = cmdF4.Top
End With
With cmdF8
    .Width = TextWidth(.Caption) + 300
    .Height = wpara.ButtonY
    .Left = cmdF6.Left + cmdF6.Width + 90
    .Top = cmdF6.Top
End With
With cmdF9
    .Width = TextWidth(.Caption) + 300
    .Height = wpara.ButtonY
    .Left = cmdF8.Left + cmdF8.Width + 90
    .Top = cmdF8.Top
End With


With cmdEsc
  .Width = wpara.ButtonX
  .Height = wpara.ButtonY
  .Left = chtPvs(4).Left + chtPvs(4).Width - .Width
  .Top = chtPvs(4).Top + chtPvs(4).Height + 90
End With

Call EinlesenAnalysen


h$ = "GrundAnalyse"
gef% = False
With cboAnalysen
    For i% = 0 To (.ListCount - 1)
        If (UCase(h$) = UCase(.List(i%))) Then
            .ListIndex = i%
            gef% = True
            Exit For
        End If
    Next i%
    If (gef% = False) And (.ListCount > 0) Then
        .ListIndex = 0
        h$ = .List(0)
    End If
    If (.ListCount = 1) Then
        Call HoleAnalyse(.text)
    End If
    If (.ListCount < 2) Then
        Call ZeigeAnalyse
    End If
End With

'Call HoleAnalyse(h$)
'Call ZeigeAnalyse

'PersonalWochenStunden% = OrgPersonalWochenStunden%
'erg% = MachAuswertung(StartDatum%, StopDatum%)

'cmdAnzDiagramme(0).Value = True

Call DefErrPop
End Sub

Sub InitChartPositions()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitChartPositions")
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
Dim i%, j%

For j% = 1 To 4
    With chtPvs(j%)
        If (j% = 1) Then
            .Width = (Me.ScaleWidth - 3 * wpara.LinksX) / 2
            .Height = (Me.ScaleHeight - 4 * wpara.TitelY - wpara.ButtonY - 2 * cmbTyp(0).Height) / 2
            .Left = wpara.LinksX
            .Top = wpara.TitelY + cmbTyp(0).Height
        ElseIf (j% = 2) Then
            .Width = chtPvs(1).Width
            .Height = chtPvs(1).Height
            .Left = chtPvs(1).Left + chtPvs(1).Width + 150
            .Top = chtPvs(1).Top
        ElseIf (j% = 3) Then
            .Width = chtPvs(1).Width
            .Height = chtPvs(1).Height
            .Left = chtPvs(1).Left
            .Top = chtPvs(1).Top + chtPvs(1).Height + 150 + cmbTyp(0).Height
        Else
            .Width = chtPvs(1).Width
            .Height = chtPvs(1).Height
            .Left = chtPvs(1).Left + chtPvs(1).Width + 150
            .Top = chtPvs(3).Top
        End If
'        .Visible = True
    End With

    With cmbTyp(j%)
      .Left = chtPvs(j%).Left
      .Top = chtPvs(j%).Top - .Height
      .Width = chtPvs(j%).Width / 3
      .TabIndex = j% * 2
      
'      .Visible = True
    End With
    
    With cmbWas(j%)
      .Left = cmbTyp(j%).Left + cmbTyp(j%).Width + 60
      .Top = cmbTyp(j%).Top
      .Width = chtPvs(j%).Width / 3
'      .Width = chtPvs(j%).Left + chtPvs(j%).Width - .Left
      .TabIndex = j% * 2 + 1
      
'      .Visible = True
    End With
    
Next j%

With chtPvs(0)
    .Width = chtPvs(2).Left + chtPvs(2).Width - chtPvs(1).Left
    .Height = chtPvs(3).Top + chtPvs(3).Height - chtPvs(1).Top
    .Left = chtPvs(1).Left
    .Top = chtPvs(1).Top
'    .Visible = True
End With
With cmbTyp(0)
  .Left = chtPvs(0).Left
  .Top = chtPvs(0).Top - .Height
  .Width = chtPvs(0).Width / 4
  .TabIndex = j% * 2
'  .Visible = True
End With
With cmbWas(0)
  .Left = cmbTyp(0).Left + cmbTyp(0).Width + 60
  .Top = cmbTyp(0).Top
  .Width = chtPvs(0).Width / 4
'  .Width = chtPvs(0).Left + chtPvs(0).Width - .Left
  .TabIndex = j% * 2 + 1
'  .Visible = True
End With
        
With flxBerechnungenGlobal
    .Rows = 4
    .Cols = 2
    .FixedRows = 0
    .FixedCols = 1
    
    .ColWidth(0) = TextWidth("Leistungspotential des Personals") + 900
    .ColWidth(1) = TextWidth("XX 99.99.99 - XX 99.99.99") + 450
    .ColAlignment(1) = flexAlignRightCenter
    
    .Width = .ColWidth(0) + .ColWidth(1) + 90
    .Height = 4 * .RowHeight(0) + 90
    .Left = chtPvs(0).Left + (chtPvs(0).Width - .Width) / 2
    
    lblAnalysen.Left = .Left
    lblAnalysen.Top = chtPvs(0).Top
    
    cboAnalysen.Left = .Left + .ColWidth(0) + 45
    cboAnalysen.Width = .ColWidth(1)
    cboAnalysen.Top = lblAnalysen.Top
    
    lblVergleichAnalyse.Left = .Left + .Width
    lblVergleichAnalyse.Width = .ColWidth(1)
    lblVergleichAnalyse.Top = lblAnalysen.Top
    
    .Top = lblAnalysen.Top + lblAnalysen.Height + 150
    
    .Visible = True
End With

For j% = 0 To 2
    With flxBerechnungen(j%)
        .Width = (Me.ScaleWidth - 3 * wpara.LinksX - 300) / 3
        If (j% = 0) Then
            .Left = chtPvs(0).Left
        Else
            .Left = flxBerechnungen(j% - 1).Left + flxBerechnungen(j% - 1).Width + 150
        End If
        
        .Rows = AnzStunden% + 1 + 1
        .FixedRows = 1
        .Cols = 7
        .FormatString = "|>Mo|>Di|>Mi|>Do|>Fr|>Sa"
        For i% = 1 To AnzStunden%
            .TextMatrix(i%, 0) = StundenBez$(i% - 1)
        Next i%
        .TextMatrix(.Rows - 1, 0) = "Summe"
        .ColWidth(0) = TextWidth("0755 - 0900")
        For i% = 1 To (.Cols - 1)
            .ColWidth(i%) = (.Width - .ColWidth(0) - 90) / 6
        Next i%
        
        .Height = .Rows * .RowHeight(0) + 90
'        .Top = chtPvs(0).Top + (chtPvs(0).Height - .Height) / 2
        .Top = flxBerechnungenGlobal.Top + flxBerechnungenGlobal.Height + 900
        .Visible = True
    End With
    
    With lblBerechnungen(j%)
        .Width = flxBerechnungen(j%).Width
        .Left = flxBerechnungen(j%).Left + 60
        .Top = flxBerechnungen(j%).Top - .Height - 30
        If (j% = 0) Then
            .Caption = "Besetzung exakt"
        ElseIf (j% = 1) Then
            .Caption = "Gerundete Besetzung"
        Else
            .Caption = "Kundenanzahl"
        End If
        .Visible = True
    End With
Next j%


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
Dim i%, ind%, h$

If (KeyCode = vbKeyF2) Then
    cmdF2.Value = True
ElseIf (KeyCode = vbKeyF3) Then
    cmdF3.Value = True
ElseIf (KeyCode = vbKeyF4) Then
    cmdF4.Value = True
ElseIf (KeyCode = vbKeyF6) Then
    cmdF6.Value = True
ElseIf (KeyCode = vbKeyF8) Then
    cmdF8.Value = True
ElseIf (KeyCode = vbKeyF9) Then
    cmdF9.Value = True
End If

Call DefErrPop
End Sub

Private Sub Form_Unload(Cancel As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Unload")
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

For i% = 1 To 4
    Unload cmbTyp(i%)
    Unload cmbWas(i%)
Next i%
For i% = 1 To 2
    Unload flxBerechnungen(i%)
    Unload lblBerechnungen(i%)
Next i%

Call DefErrPop
End Sub

Private Sub cmdF6_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF6_Click")
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

If (cmdF6.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

Call DruckeWindow(Me)
If (AnzeigeModus% = 0) And (flxBerechnungenGlobal.Cols = 2) Then
    Call DruckePlanung
End If

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
Dim i%, ind%
Dim h$, ab$, Bis$

If (cmdF2.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

h$ = flxBerechnungenGlobal.TextMatrix(0, 1)
Do
    ind% = InStr(h$, ".")
    If (ind% > 0) Then
        h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
    Else
        Exit Do
    End If
Loop
ind% = InStr(h$, "-")
ab$ = Trim(Left$(h$, ind% - 1))
Bis$ = Trim(Mid$(h$, ind% + 1))
        

For i% = 1 To 2
    Load txtEditAnalyse(i%)
Next i%
For i% = 0 To 2
    With txtEditAnalyse(i%)
        .BackColor = vbGreen
        .Width = flxBerechnungenGlobal.ColWidth(1) / 2
        .Visible = True
        .ZOrder 0
    End With
Next i%
With txtEditAnalyse(0)
    .Left = flxBerechnungenGlobal.Left + flxBerechnungenGlobal.ColPos(1) + 45
    .Top = flxBerechnungenGlobal.Top + 45
'    .Text = ab$
    .text = sDate(OrgStartDatum%)
    .SetFocus
End With
With txtEditAnalyse(1)
    .Left = txtEditAnalyse(0).Left + txtEditAnalyse(0).Width + 15
    .Top = flxBerechnungenGlobal.Top + 45
'    .Text = Bis$
    .text = sDate(OrgStopDatum%)
End With
With txtEditAnalyse(2)
    .Left = txtEditAnalyse(1).Left
    .Top = flxBerechnungenGlobal.Top + flxBerechnungenGlobal.RowHeight(0) + 45
'    .Text = flxBerechnungenGlobal.TextMatrix(1, 1)
    .text = Format(OrgPersonalWochenStunden%)
End With

cmdF2.Enabled = False
cmdF3.Enabled = False
cmdF4.Enabled = False
cmdF6.Enabled = False
cmdF8.Enabled = False
cmdF9.Enabled = False
cboAnalysen.Enabled = False

Call DefErrPop
End Sub

Private Sub cmdF3_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF3_Click")
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
Dim i%, ind%
Dim h$, ab$, Bis$

If (cmdF3.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

h$ = flxBerechnungenGlobal.TextMatrix(0, 1)
Do
    ind% = InStr(h$, ".")
    If (ind% > 0) Then
        h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
    Else
        Exit Do
    End If
Loop
ind% = InStr(h$, "-")
ab$ = Trim(Left$(h$, ind% - 1))
Bis$ = Trim(Mid$(h$, ind% + 1))
        

For i% = 1 To 2
    Load txtEditAnalyse(i%)
Next i%
For i% = 0 To 2
    With txtEditAnalyse(i%)
        .BackColor = vbGreen
        .Width = flxBerechnungenGlobal.ColWidth(1) / 2
        .Visible = True
        .ZOrder 0
    End With
Next i%
With txtEditAnalyse(0)
    .Left = flxBerechnungenGlobal.Left + flxBerechnungenGlobal.ColPos(1) + 45
    .Top = flxBerechnungenGlobal.Top + 45
'    .Text = ab$
    .text = sDate(OrgStartDatum%)
    .SetFocus
End With
With txtEditAnalyse(1)
    .Left = txtEditAnalyse(0).Left + txtEditAnalyse(0).Width + 15
    .Top = flxBerechnungenGlobal.Top + 45
'    .Text = Bis$
    .text = sDate(OrgStopDatum%)
End With
With txtEditAnalyse(2)
    .Left = txtEditAnalyse(1).Left
    .Top = flxBerechnungenGlobal.Top + flxBerechnungenGlobal.RowPos(3) + 45
'    .Text = flxBerechnungenGlobal.TextMatrix(1, 1)
    .text = flxBerechnungenGlobal.TextMatrix(3, 1)
End With

cmdF2.Enabled = False
cmdF3.Enabled = False
cmdF4.Enabled = False
cmdF6.Enabled = False
cmdF8.Enabled = False
cmdF9.Enabled = False
cboAnalysen.Enabled = False

PbaTest% = True

Call DefErrPop
End Sub

Private Sub cmdF8_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF8_Click")
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

If (cmdF8.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

PbaWahlModus% = 1
frmPbaSave.Show 1

If (EditErg%) Then
    Vergleich% = True
    For i% = 0 To UBound(PbaRec)
        PbaRec2(i%) = PbaRec(i%)
    Next i%
    
    lblVergleichAnalyse.Caption = EditTxt$
    
    If (EditTxt$ = "Neuer Pba-Analyse") Then
        cmdF2.Value = True
    Else
        Call HoleAnalyse(EditTxt$)
        Call ZeigeAnalyse
    End If
End If

Call DefErrPop
End Sub

Private Sub cmdF9_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF9_Click")
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

If (cmdF9.Enabled = False) Then
    Call DefErrPop: Exit Sub
End If

PbaWahlModus% = 0
frmPbaSave.Show 1

If (EditErg%) Then
    Call SpeicherAnalyse(EditTxt$)
End If

Call DefErrPop
End Sub

Private Sub txtEditAnalyse_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEditAnalyse_GotFocus")
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
Dim h$

With txtEditAnalyse(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtEditAnalyse_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEditAnalyse_KeyPress")
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
Dim i%, erg%

If (KeyAscii = vbKeyReturn) Then
    If (Index = 2) Then
        If (PbaTest%) Then
            KundenProArbeitsStunde# = Val(txtEditAnalyse(2).text) / NORMTAGZEIT% * 60#
        Else
            PersonalWochenStunden% = Val(txtEditAnalyse(2).text)
        End If
        StartDatum% = iDate(txtEditAnalyse(0).text)
        StopDatum% = iDate(txtEditAnalyse(1).text)
        For i% = 1 To 2
            Unload txtEditAnalyse(i%)
        Next i%
        txtEditAnalyse(0).Visible = False
    
        cmdF2.Enabled = True
        cmdF3.Enabled = True
        cmdF4.Enabled = True
        cmdF6.Enabled = True
        cmdF8.Enabled = True
        cmdF9.Enabled = True
        cboAnalysen.Enabled = True
        cboAnalysen.SetFocus
        
        erg% = MachAuswertung(StartDatum%, StopDatum%)
        If (erg%) Then
            Call ZeigeAnalyse
        End If
        PbaTest% = False
    Else
        txtEditAnalyse(Index + 1).SetFocus
    End If
End If

Call DefErrPop
End Sub

Sub ZeigeAnalyse()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeAnalyse")
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
Dim i%, j%, k%, r%, g%, b%, col%, diff%, WoTag%
Dim LegendenPos&
Dim ZeitStr$, h$, ab$, Bis$, a$, b2$
Dim Reihe As Object
Dim punkt As Object
Dim iStundenPers%, iStundenPers2%
Dim KundenProStunde#, PersonalProStunde#, dStundenPers2#, SummeTag#
Dim WochenTag$(6)

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"

For i% = 0 To 2
    With lblBerechnungen(i%)
        If (Vergleich%) Then
            If (i% = 0) Then
                .Caption = "Differenz Sollb. in %"
            ElseIf (i% = 1) Then
                .Caption = "Differenz Sollb. gerund."
            Else
                .Caption = "Differenz Kunden in %"
            End If
        Else
            If (i% = 0) Then
                .Caption = "Besetzung exakt"
            ElseIf (i% = 1) Then
                .Caption = "Gerundete Besetzung"
            Else
                .Caption = "Kundenanzahl"
            End If
        End If
    End With
Next i%

For k% = 0 To 2
    With flxBerechnungen(k%)
         For j% = 1 To 6
            col% = 1
            For i% = 0 To 23
                KundenProStunde# = PbaRec(j%).KundenProStunde(i%)
                If (KundenProStunde# > 0) Then
                    PersonalProStunde# = PbaRec(j%).PersonalProStunde#(i%)
                    dStundenPers2# = PbaRec2(j%).PersonalProStunde#(i%)
                    
                    iStundenPers% = PbaRec(j%).iPersonalProStunde(i%)
                    iStundenPers2% = PbaRec2(j%).iPersonalProStunde(i%)
                Else
                    PersonalProStunde# = 0
                    dStundenPers2# = 0
                    
                    iStundenPers% = 0
                    iStundenPers2% = 0
                End If
                
                h$ = ""
                If (KundenProStunde# > 0#) Then
                    If (Vergleich%) Then
                        If (k% = 0) Then
                            diff% = CInt((dStundenPers2# / PersonalProStunde#) * 100) - 100
                            If (dStundenPers2# = 0#) Then
                                h$ = "kA"
                            ElseIf (Abs(diff%) <= ToleranzVergleich%) Then
                                h$ = ""
                            Else
                                h$ = Format(diff%, "0")
                            End If
                        ElseIf (k% = 1) Then
                            diff% = iStundenPers2% - iStundenPers%
                            If (diff% = 0) Then
                                h$ = ""
                            Else
                                h$ = Format(diff%, "0")
                            End If
                        Else
                            diff% = CInt((PbaRec2(j%).KundenProStunde(i%) / KundenProStunde#) * 100) - 100
                            If (PbaRec2(j%).KundenProStunde(i%) = 0#) Then
                                h$ = "kA"
                            ElseIf (Abs(diff%) <= ToleranzVergleich%) Then
                                h$ = ""
                            Else
                                h$ = Format(diff%, "0")
                            End If
                        End If
                    Else
                        If (k% = 0) Then
                            h$ = Format(PersonalProStunde#, "0.0")
                        ElseIf (k% = 1) Then
                            h$ = Format(iStundenPers%, "0")
                        Else
                            h$ = Format(KundenProStunde#, "0")
                        End If
                    End If
                End If
                If (col% <= AnzStunden%) Then
                    .TextMatrix(col%, j%) = h$
                End If
                    
                If (i% >= gVon%) And (i% <= gBis%) Then
                    col% = col% + 1
                End If
            Next i%
        
            SummeTag# = PbaRec(j%).KundenProTag
            h$ = ""
            If (Vergleich% = 0) And (SummeTag# > 0) Then
                If (k% = 0) Then
                    h$ = Format(PbaRec(j%).PersonalProTag, "0.0")
                ElseIf (k% = 1) Then
                    h$ = Format(PbaRec(j%).iPersonalProTag, "0")
                Else
                    h$ = Format(SummeTag#, "0")
                End If
            End If
            .TextMatrix(.Rows - 1, j%) = h$
        Next j%
        
        .FillStyle = flexFillRepeat
        .row = .Rows - 1
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellFontBold = True
        .FillStyle = flexFillSingle
    End With
Next k%
             
With flxBerechnungenGlobal
    ab$ = sDate(StartDatum%)
    h$ = Left$(ab$, 2) + "." + Mid$(ab$, 3, 2) + "." + Mid$(ab$, 5, 2)
    WoTag% = WeekDay(h$, vbMonday) - 1
    ZeitStr$ = Left$(WochenTag$(WoTag%), 2) + " " + h$
    ZeitStr$ = ZeitStr$ + "  -  "
    
    Bis$ = sDate(StopDatum%)
    h$ = Left$(Bis$, 2) + "." + Mid$(Bis$, 3, 2) + "." + Mid$(Bis$, 5, 2)
    WoTag% = WeekDay(h$, vbMonday) - 1
    ZeitStr$ = ZeitStr$ + Left$(WochenTag$(WoTag%), 2) + " " + h$
        
    
    
    If (Vergleich%) Then
        .Cols = 3
        .ColWidth(2) = .ColWidth(1)
        .ColAlignment(2) = flexAlignRightCenter
        .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
        
        .TextMatrix(0, 2) = ZeitStr$
        .TextMatrix(1, 2) = Format(PersonalWochenStunden%, "0")
        .TextMatrix(2, 2) = Format(KundenProTag#, "0")
    '    .TextMatrix(3, 2) = Format(KundenProArbeitsStunde#, "0")
        .TextMatrix(3, 2) = Format(KundenProArbeitsStunde# * (NORMTAGZEIT% / 60#), "0")
        
        For i% = 0 To UBound(PbaRec)
            PbaRec(i%) = PbaRec2(i%)
        Next i%
    Else
        .Cols = 2
        .Width = .ColWidth(0) + .ColWidth(1) + 90
        
        .TextMatrix(0, 0) = "Beobachtungs-Zeitraum"
        .TextMatrix(2, 0) = "Kundenanzahl / Tag"
        If (PbaTest%) Then
            .TextMatrix(1, 0) = "Personalstunden / Woche  (GEPLANT)"
            .TextMatrix(3, 0) = "Kunden / Normtag   (GEPLANT)"
        Else
            .TextMatrix(1, 0) = "Personalstunden / Woche"
            .TextMatrix(3, 0) = "Kunden / Normtag"
        End If
        .TextMatrix(0, 1) = ZeitStr$
        .TextMatrix(1, 1) = Format(PersonalWochenStunden%, "0")
        .TextMatrix(2, 1) = Format(KundenProTag#, "0")
    '    .TextMatrix(3, 1) = Format(KundenProArbeitsStunde#, "0")
        .TextMatrix(3, 1) = Format(KundenProArbeitsStunde# * (NORMTAGZEIT% / 60#), "0")
    End If
End With

lblVergleichAnalyse.Visible = Vergleich%
cmdF4.Enabled = (Vergleich% = 0)

For i% = 0 To 4
    Call cmbWas_Click(i%)
Next i%

Vergleich% = False

Call DefErrPop
End Sub

Private Sub cboAnalysen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cboAnalysen_Click")
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

If (cboAnalysen.Enabled) Then
    Call HoleAnalyse(cboAnalysen.text)
    Call ZeigeAnalyse
End If

Call DefErrPop
End Sub

Sub DruckePlanung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckePlanung")
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
Dim i%, j%, k%, einh%, pers%(24), x%, Y%, leer%
Dim l&
Dim h$, heute$
Dim WochenTag$(6)

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"

With flxBerechnungen(1)
    einh% = (Printer.ScaleWidth - 450) / (.Rows - 2)
    
    heute$ = Format(Day(Date), "00") + "-"
    heute$ = heute$ + Format(Month(Date), "00") + "-"
    heute$ = heute$ + Format(Year(Date), "0000")

    
    For i% = 1 To 6
        If (i% = 1) Or (i% = 4) Then
            Printer.CurrentX = 0
            Printer.CurrentY = 0
            Printer.Font.Size = 18
            Printer.Print "Personalbedarf-Analyse: Planungshilfe";
            
            l& = Printer.TextWidth(heute$)
            Printer.CurrentX = Printer.ScaleWidth - l& - 150
            Printer.CurrentY = 0
            Printer.Print heute$
            Printer.Print
        End If
        
        Printer.CurrentX = 0
        Printer.Font.Name = "Arial"
        Printer.Font.Size = 16
        Printer.Print WochenTag$(i% - 1)
        
        Printer.CurrentY = Printer.CurrentY + 150
        
        Printer.Font.Size = 12
        For j% = 1 To (.Rows - 2)
            pers%(j%) = Val(.TextMatrix(j%, i%))
            Printer.CurrentX = (j% - 1) * einh% + 45
            Printer.Print Mid$(.TextMatrix(j%, 0), 6, 2);
        Next j%
        Printer.Print
        Y% = Printer.CurrentY
        
        Do
            Printer.CurrentX = 0
            x% = Printer.CurrentX
            Y% = Printer.CurrentY
            leer% = True
            For j% = 1 To (.Rows - 2)
                If (pers%(j%) > 0) Then
                    Printer.Line (x%, Y%)-(x% + einh%, Y% + 600), vbBlack, B
                    pers%(j%) = pers%(j%) - 1
                    leer% = 0
                End If
                x% = x% + einh%
            Next j%
            If (leer%) Then Exit Do
            Printer.CurrentY = Y% + 600
        Loop

        Printer.CurrentY = Printer.CurrentY + 1050
        
        If (i% = 3) Then Printer.NewPage
    Next i%
End With

Printer.EndDoc


Call DefErrPop
End Sub



