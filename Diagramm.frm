VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmDiagramm 
   Caption         =   "Diagramme"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   570
   ClientWidth     =   9660
   Icon            =   "Diagramm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   Begin VB.PictureBox picAusdruck 
      Height          =   735
      Left            =   6240
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdF6 
      Caption         =   "&Drucken (F6)"
      Height          =   975
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdAnzDiagramme 
      Caption         =   "&4 Diagramme gleichzeitig"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdAnzDiagramme 
      Caption         =   "Nur &1 Diagramm "
      Enabled         =   0   'False
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.ComboBox cmbWas 
      Height          =   315
      Index           =   0
      Left            =   8760
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   6000
      Width           =   2175
   End
   Begin VB.ComboBox cmbTyp 
      Height          =   315
      Index           =   0
      Left            =   6720
      Style           =   2  'Dropdown-Liste
      TabIndex        =   0
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdKreis 
      Caption         =   "&Kreisdiagramm"
      Height          =   975
      Left            =   8520
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdBalken 
      Caption         =   "&Balkendiagramm"
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Esc"
      Height          =   975
      Left            =   10680
      TabIndex        =   6
      Top             =   6000
      Width           =   2175
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   2
      Left            =   -120
      OleObjectBlob   =   "Diagramm.frx":030A
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   1
      Left            =   7560
      OleObjectBlob   =   "Diagramm.frx":27D5
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   4575
      Index           =   0
      Left            =   360
      OleObjectBlob   =   "Diagramm.frx":4CA0
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   3
      Left            =   3600
      OleObjectBlob   =   "Diagramm.frx":716B
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSChartLib.MSChart chtPvs 
      Height          =   5535
      Index           =   4
      Left            =   4920
      OleObjectBlob   =   "Diagramm.frx":9636
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmDiagramm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpInd%(50)
Dim ActDiagrammInd%
Dim AnzSp%
Private Const DefErrModul = "Diagramm.FRM"

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

DiagrammTyp%(Index) = cmbTyp(Index).ListIndex
ActDiagrammInd% = Index
Call SpeicherIniDiagramme(Index)
  
With chtPvs(Index).DataGrid
    .ColumnCount = AnzSp%
    If (typ = VtChChartType2dBar) Then
        ind% = cmbWas(Index).ListIndex
'        If (InStr(cmbWas(Index).Text, "/") > 0) Then
        If (ind% >= 0) Then
            If (dInfo$(ind%, 2, 0) = "#") Then
                .ColumnCount = AnzSp% + 1
                .ColumnLabel(AnzSp% + 1, 1) = "ges."
                .SetData 1, AnzSp% + 1, xVal(dInfo$(ind%, 1, 0)), False
                chtPvs(Index).Plot.SeriesCollection(AnzSp% + 1).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
            End If
        End If
    End If
End With
Call DefErrPop
End Sub

Private Sub cmbTyp_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbTyp_GotFocus")
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
ActDiagrammInd% = Index
Call DefErrPop
End Sub

Private Sub cmbWas_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbWas_GotFocus")
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
ActDiagrammInd% = Index
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
Dim ind%, i%
  
With chtPvs(Index)
    ind% = cmbWas(Index).ListIndex
    With .DataGrid
        .ColumnCount = AnzSp%
        .RowLabel(1, 1) = ""   ' dInfo$(ind%, 0, 0)
      
        For i% = 1 To .ColumnCount
        '    .SetData 1, i%, CInt(xVal(dInfo$(ind%, 1, SpInd%(i% - 1)))), False 'geht nichtwegen Tausender-Trennzeichen
            If (Val(dInfo$(ind%, 1, SpInd%(i% - 1))) = 0) Then
                .SetData 1, i%, 0, False
            Else
                .SetData 1, i%, CInt(dInfo$(ind%, 1, SpInd%(i% - 1))), False
            End If
        Next i%
    
        If (cmbTyp(Index).ListIndex = 0) Then
'            If (InStr(cmbWas(Index).Text, "/") > 0) Then
            If (dInfo$(ind%, 2, 0) = "#") Then
                .ColumnCount = AnzSp% + 1
                .ColumnLabel(AnzSp% + 1, 1) = "ges."
                .SetData 1, AnzSp% + 1, xVal(dInfo$(ind%, 1, 0)), False
                chtPvs(Index).Plot.SeriesCollection(AnzSp% + 1).DataPoints(-1).Brush.FillColor.Set 0, 0, 0
            End If
        End If

    End With
    DiagrammWas%(Index) = ind%
    ActDiagrammInd% = Index
    Call SpeicherIniDiagramme(Index)
End With
Call DefErrPop

End Sub

Private Sub cmdAnzDiagramme_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAnzDiagramme_Click")
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

With chtPvs(ActDiagrammInd%)
    If (Index = 0) Then
        For i% = 1 To 4
            chtPvs(i%).Visible = False
            cmbTyp(i%).Visible = False
            cmbWas(i%).Visible = False
        Next i%
        chtPvs(0).Visible = True
        cmbTyp(0).Visible = True
        cmbWas(0).Visible = True
        cmbTyp(0).SetFocus
        
        cmdAnzDiagramme(0).Enabled = False
        cmdAnzDiagramme(1).Enabled = True
    Else
        For i% = 1 To 4
            chtPvs(i%).Visible = True
            cmbTyp(i%).Visible = True
            cmbWas(i%).Visible = True
        Next i%
        chtPvs(0).Visible = False
        cmbTyp(0).Visible = False
        cmbWas(0).Visible = False
        cmbTyp(1).SetFocus
        
        cmdAnzDiagramme(0).Enabled = True
        cmdAnzDiagramme(1).Enabled = False
    End If
End With
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

Call DruckeWindow(Me)

Call DefErrPop
End Sub

'Private Sub cmdF6_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdF6_Click")
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
'Dim border, faktor
'Dim i%, MaxWi%, MaxHe%, wi%, he%, TextHe%, CurrX%, CurrY%
'Dim h$
'
'With Printer
'    .Orientation = vbPRORLandscape
'    .ScaleMode = vbTwips
'    .Font.Name = "Arial"
'
'    border = .ScaleWidth / 20
'
'    If (chtPvs(0).Visible) Then
'        .Font.Size = 18
'        faktor = chtPvs(0).Width / chtPvs(0).Height
'        TextHe% = .TextHeight("Äg")
'        MaxWi% = .ScaleWidth - 2 * border
'        MaxHe% = .ScaleHeight - 2 * border - TextHe%
'    Else
'        .Font.Size = 14
'        faktor = chtPvs(1).Width / chtPvs(1).Height
'        TextHe% = .TextHeight("Äg")
'        MaxWi% = (.ScaleWidth - 2 * border - 300) / 2
'        MaxHe% = (.ScaleHeight - 2 * (border + TextHe%) - 300) / 2
'    End If
'
'    wi% = MaxWi%
'    he% = wi% / faktor
'    If (he% > MaxHe%) Then
'        he% = MaxHe%
'        wi% = he% * faktor
'    End If
'
'    If (chtPvs(0).Visible) Then
'        h$ = cmbWas(0).Text
'        .CurrentX = border + (MaxWi% - .TextWidth(h$)) / 2
'        .CurrentY = border
'        Printer.Print h$
'
'        chtPvs(0).EditCopy
'        Set picAusdruck.Picture = Clipboard.GetData(vbCFMetafile)
'        Printer.PaintPicture picAusdruck.Picture, border, border + TextHe%, wi%, he%
'    Else
'        For i% = 1 To 4
'            h$ = cmbWas(i%).Text
'            chtPvs(i%).EditCopy
'            Set picAusdruck.Picture = Clipboard.GetData(vbCFMetafile)
'
'            If (i% = 1) Or (i% = 3) Then
'                CurrX% = border + (MaxWi% - .TextWidth(h$)) / 2
'            Else
'                CurrX% = border + MaxWi% + 300 + (MaxWi% - .TextWidth(h$)) / 2
'            End If
'            If (i% <= 2) Then
'                CurrY% = border
'            Else
'                CurrY% = border + MaxHe% + 300
'            End If
'            .CurrentX = CurrX%
'            .CurrentY = CurrY%
'            Printer.Print h$
'
'            If (i% = 1) Or (i% = 3) Then
'                CurrX% = border
'            Else
'                CurrX% = border + MaxWi% + 300
'            End If
'            If (i% <= 2) Then
'                CurrY% = border
'            Else
'                CurrY% = border + MaxHe% + 300
'            End If
'            Printer.PaintPicture picAusdruck.Picture, CurrX%, CurrY% + TextHe%, wi%, he%
'        Next i%
'    End If
'
'    .EndDoc
'End With
'
'
'
'
''If (cmdAnzDiagramme(0).Enabled) Then
''Else
''End If
'
'Call DefErrPop
'End Function

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

Dim i%, j%, r%, g%, b%, ind%
Dim LegendenPos&
Dim h$
Dim Reihe As Object
Dim punkt As Object

Call wpara.InitFont(Me)

Me.Width = frmAction.Width - 2 * wpara.FrmCaptionHeight
Me.Height = frmAction.Height - 2 * wpara.FrmCaptionHeight
Me.Left = frmAction.Left + wpara.FrmCaptionHeight
Me.Top = frmAction.Top + wpara.FrmCaptionHeight

For i% = 1 To 4
    Load cmbTyp(i%)
    Load cmbWas(i%)
Next i%

On Error Resume Next
AnzSp% = 0

If Vorschau Then
  For j% = 1 To UBound(DiagPnr)
    ind% = DiagPnr%(j%)
    If (Val(dInfo$(HAT_WERTE_SPALTE%, 1, ind%)) > 0) Then
      SpInd%(AnzSp%) = ind%
      AnzSp% = AnzSp% + 1
    End If
    
  Next j%
Else
  With frmAction!cboMitarbeiter
  
    For j% = 1 To .ListCount - 1
      For i% = 1 To UBound(MitArb$)
        If Trim(Left(MitArb$(i%), 20)) = .List(j%) Then
          ind% = Val(Mid(MitArb$(i%), 22, 2))
          Exit For
        End If
      Next i%
    'For i% = 1 To UBound(dInfo$(), 3)
      'If (Val(dInfo$(i% - 1, 1, i%)) > 0) Then
      'Kd/NT
      If (Val(dInfo$(HAT_WERTE_SPALTE%, 1, ind%)) > 0) Then
        SpInd%(AnzSp%) = ind%
        AnzSp% = AnzSp% + 1
      End If
    Next j%
  End With
End If
On Error GoTo 0

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
            .ColumnCount = AnzSp%
            
            For i% = 1 To .ColumnCount
                If (PersonalInitialen$(SpInd%(i% - 1)) <> "") Then
                    .ColumnLabel(i%, 1) = PersonalInitialen$(SpInd%(i% - 1))
                Else
                    .ColumnLabel(i%, 1) = para.Personal(SpInd%(i% - 1))
                End If
'                .ColumnLabel(i%, 1) = Left$(para.Personal(SpInd%(i% - 1)), 2)
            Next i%
            
            .RowCount = 1
            
            For i% = 1 To chtPvs(j%).Plot.SeriesCollection.Count
                With chtPvs(j%).Plot.SeriesCollection(i%).DataPoints(-1)
                    h$ = PersonalFarben$(SpInd%(i% - 1))
                    b% = Val("&H" + Left$(h$, 2))
                    g% = Val("&H" + Mid$(h$, 3, 2))
                    r% = Val("&H" + Mid$(h$, 5, 2))
                    .Brush.FillColor.Set r%, g%, b%
'                    .Brush.FillColor.Set Val(d"&H" + Left$(PersonalFarben$(SpInd%(i% - 1)), 2)), Val("&H" + Mid$(PersonalFarben$(SpInd%(i% - 1)), 3, 2)), Val("&H" + Mid$(PersonalFarben$(SpInd%(i% - 1)), 5, 2))
                End With
'                chtPvs(j%).Plot.SeriesCollection(i%).DataPoints(0).Marker
            Next
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
        
        .ListIndex = DiagrammTyp%(i%)
    End With
    
    With cmbWas(i%)
        For j% = 0 To UBound(dInfo$(), 1)
            .AddItem dInfo$(j%, 0, 0)
        Next j%
      
        .ListIndex = DiagrammWas%(i%)
    End With
Next i%
ActDiagrammInd% = 0


For i% = 0 To 1
    With cmdAnzDiagramme(i%)
        .Width = TextWidth(.Caption) + 150
        .Height = wpara.ButtonY
        If (i% = 0) Then
            .Left = chtPvs(0).Left
        Else
            .Left = cmdAnzDiagramme(i% - 1).Left + cmdAnzDiagramme(i% - 1).Width + 150
        End If
        .Top = chtPvs(3).Top + chtPvs(3).Height + 90
    End With
Next i%

With cmdF6
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = cmdAnzDiagramme(1).Left + cmdAnzDiagramme(1).Width + 450
    .Top = cmdAnzDiagramme(1).Top
End With

With cmdEsc
  .Width = wpara.ButtonX
  .Height = wpara.ButtonY
  .Left = chtPvs(4).Left + chtPvs(4).Width - .Width
  .Top = chtPvs(4).Top + chtPvs(4).Height + 90
End With

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
Dim j%

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
      .Width = chtPvs(j%).Left + chtPvs(j%).Width - .Left
      .TabIndex = j% * 2 + 1
      
'      .Visible = True
    End With
Next j%

With chtPvs(0)
    .Width = chtPvs(2).Left + chtPvs(2).Width - chtPvs(1).Left
    .Height = chtPvs(3).Top + chtPvs(3).Height - chtPvs(1).Top
    .Left = chtPvs(1).Left
    .Top = chtPvs(1).Top
    .Visible = True
End With
With cmbTyp(0)
  .Left = chtPvs(0).Left
  .Top = chtPvs(0).Top - .Height
  .Width = chtPvs(0).Width / 3
  .TabIndex = j% * 2
  .Visible = True
End With
With cmbWas(0)
  .Left = cmbTyp(0).Left + cmbTyp(0).Width + 60
  .Top = cmbTyp(0).Top
  .Width = chtPvs(0).Left + chtPvs(0).Width - .Left
  .TabIndex = j% * 2 + 1
  .Visible = True
End With
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

If (KeyCode = vbKeyF6) Then
    cmdF6.Value = True
End If

Call DefErrPop
End Sub


