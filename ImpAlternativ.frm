VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmImpAlternativ 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Import-Alternativen"
   ClientHeight    =   4680
   ClientLeft      =   585
   ClientTop       =   1155
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9855
   Begin VB.PictureBox picAnimationBack 
      Appearance      =   0  '2D
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   1200
      ScaleHeight     =   2370
      ScaleWidth      =   5625
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   5655
      Begin ComCtl2.Animation aniAnimation 
         Height          =   1095
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1931
         _Version        =   327681
         Center          =   -1  'True
         BackColor       =   -2147483624
         FullWidth       =   73
         FullHeight      =   73
      End
      Begin VB.Label lblAnimation 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         Caption         =   "Aufgabe wird bearbeitet ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   -240
         Width           =   5355
      End
   End
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   6600
      Top             =   360
   End
   Begin VB.CommandButton cmdF3 
      Caption         =   "Sortierung (F3)"
      Enabled         =   0   'False
      Height          =   450
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdF6 
      Caption         =   "Drucken (F6)"
      Enabled         =   0   'False
      Height          =   450
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   5640
      TabIndex        =   4
      Top             =   2040
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrg 
      Height          =   1260
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   2223
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
   Begin MSFlexGridLib.MSFlexGrid flxImporte 
      Height          =   1260
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   2223
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
End
Attribute VB_Name = "frmImpAlternativ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortCol%
Dim InSuche%, AbbruchSuche%

Private Const DefErrModul = "IMPALTERNATIV.FRM"

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

If (InSuche%) Then
    AbbruchSuche% = True
Else
    Unload Me
End If

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
Dim i%

If (cmdF3.Enabled) Then
    If (SortCol% = 1) Then
        SortCol% = 7
    ElseIf (SortCol% = 7) Then
        SortCol% = 8
    Else
        SortCol% = 1
    End If
    Call SortFlex
End If

Call DefErrPop
End Sub

Private Sub SortFlex()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SortFlex")
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

With flxOrg
    .Redraw = False
    .row = 1
    .col = SortCol%
    .RowSel = .Rows - 1
    If (SortCol% = 1) Then
        .ColSel = 4
        .Sort = 5
    Else
        .ColSel = .col
        .Sort = 4
    End If
    
    .FillStyle = flexFillRepeat
    .row = 0
    For i% = 0 To (.Cols - 1)
        .col = i%
        .CellFontBold = False
    Next i%
    .col = SortCol%
    .RowSel = .row
    If (SortCol% = 1) Then
        .ColSel = 4
    Else
        .ColSel = .col
    End If
    .CellFontBold = True
    .FillStyle = flexFillSingle
    
    .TopRow = 1
    .Redraw = True
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .SetFocus
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

If (cmdF6.Enabled) Then
    Call ImpAlternativAusdruck
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
Dim i%, Breite%, MaxWi%, wi%, diff%, FormVersatzY%
Dim c As Object

Call wpara.InitFont(Me)

FormVersatzY% = wpara.FrmCaptionHeight + wpara.FrmBorderHeight
Me.Width = frmRezSpeicher.Width
Me.Height = frmRezSpeicher.Height - FormVersatzY%
Me.Left = frmRezSpeicher.Left
Me.Top = frmRezSpeicher.Top + FormVersatzY%


''''''''''''''''''''''''
With flxOrg
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Pzn|<Name der ORIGINALE|>Menge|<Eh|<Herst|>POS|>VK|>Anz|>Ges.Wert|"
    .Rows = 1
'    .Cols = 7
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = (Me.ScaleHeight - .Top - wpara.ButtonY - 300) * 0.7
    .Height = (.Height \ .RowHeight(0)) * .RowHeight(0) + 90
    .Width = ScaleWidth - 2 * .Left

    .ColWidth(0) = TextWidth(String(11, "9"))
    .ColWidth(1) = 0 ' TextWidth(String(30, "X"))
    .ColWidth(2) = TextWidth(String(8, "X"))
    .ColWidth(3) = TextWidth(String(5, "X"))
    .ColWidth(4) = TextWidth(String(8, "X"))
    .ColWidth(5) = TextWidth(String(6, "9"))
    .ColWidth(6) = TextWidth("99999999.99")
    .ColWidth(7) = TextWidth(String(6, "9"))
    .ColWidth(8) = TextWidth("99999999.99")
    .ColWidth(9) = wpara.FrmScrollHeight
    
    If (ImpAlternativModus% = 1) Then
        .ColWidth(7) = 0
        .ColWidth(8) = 0
    End If
     
    Breite% = 0
    For i% = 0 To (.Cols - 1)
        Breite% = Breite% + .ColWidth(i%)
    Next i%
    Breite% = .Width - Breite% - 90
    If (Breite% > 0) Then
        .ColWidth(1) = Breite%
    End If
    
    .BackColor = vbWhite
End With

With flxImporte
    .Rows = 2
    .FixedRows = 1
    .FormatString = flxOrg.FormatString
    .TextMatrix(0, 1) = "Name der zugehörigen IMPORTE"
    .Rows = 1
    
    .Top = flxOrg.Top + flxOrg.Height + 150
    .Left = flxOrg.Left
    .Height = (Me.ScaleHeight - .Top - wpara.ButtonY - 300)
    .Height = (.Height \ .RowHeight(0)) * .RowHeight(0) + 90
    .Width = flxOrg.Width
    
    For i% = 0 To (.Cols - 1)
        .ColWidth(i%) = flxOrg.ColWidth(i%)
    Next i%
    .TextMatrix(0, .Cols - 2) = ""
    .TextMatrix(0, .Cols - 3) = ""
End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With cmdEsc
    .Top = flxImporte.Top + flxImporte.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
'    .Left = (ScaleWidth - .Width) / 2
    .Left = flxImporte.Left + flxImporte.Width - .Width
End With

With cmdF3
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = flxOrg.Left
    .Top = cmdEsc.Top
End With
With cmdF6
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = cmdF3.Left + cmdF3.Width + 150
    .Top = cmdEsc.Top
End With

If (ImpAlternativModus% = 0) Then
    SortCol% = 8
Else
    SortCol% = 1
End If

Call InitAnimation

Call DefErrPop
End Sub

Private Sub tmrStart_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrStart_Timer")
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
Dim AbAnzPackungen%
Dim AbGesamtWert#
Dim AbDatum$

tmrStart.Enabled = False
If (ImpAlternativModus% = 0) Then
    AbDatum$ = ImpAlternativPara$(0)
    AbAnzPackungen% = Val(ImpAlternativPara$(1))
    AbGesamtWert# = xVal(ImpAlternativPara$(2))
    Call Suche(AbDatum$, AbAnzPackungen%, AbGesamtWert#)
Else
    Call LagerOriginal
End If

Call DefErrPop
End Sub

Sub Suche(AbDatum$, Anzahl%, Wert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Suche")
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
Dim xc$, SQLStr$, h$, POSStr$
Dim found&, Max&, l&, j&
Dim IstMulti As Boolean
Dim NextMulti%, i%, LoeschFlag%
Dim AVP#, aktavp#

Call StartAnimation(Me, "Abg. Originale werden gesucht ...")

InSuche% = True
AbbruchSuche% = False
MousePointer = vbHourglass

flxOrg.Rows = 1
flxOrg.Redraw = False

xc$ = MKDate(iDate(AbDatum$))

Call vk.GetRecord(1)
Max& = vk.erstmax

found& = vk.DatumSuche(xc$)
If found& < 0 Then found& = Abs(found&) + 1
Max& = vk.DateiLen / vk.RecordLen
'TaxeRec.Index = "PZN"
j& = 1
Do While found& <= Max&
    vk.GetRecord (found&)
    IstMulti = False
    If vk.pzn = "9999999" And Mid$(vk.text$, 20, 1) = "x" Then IstMulti = True
    NextMulti% = 1
    If found& < Max& Then
        vk.GetRecord (found& + 1)
        If vk.pzn = "9999999" And Mid$(vk.text$, 20, 1) = "x" Then
            NextMulti% = Val(Mid$(vk.text$, 16, 3))
        End If
        vk.GetRecord (found&)
    End If
    If Val(vk.RezEan) > 0 And vk.pzn <> "9999999" And (Not IstMulti) Then
      If IstOriginal(Val(vk.pzn), 0) Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + vk.pzn$
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If Not TaxeRec.EOF Then
'            If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then    'GS: muss kein Original sein, um als solches zu gelten
                POSStr$ = ""
                
                If (ArtikelDbOk) Then
                    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
                '                    SQLStr = SQLStr + " AND LagerKz<>0"
                    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                Else
                    FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ass.GetRecord (FabsRecno& + 1)
                    End If
                End If
                If (FabsErrf% = 0) Then
                    POSStr$ = Format(ass.poslag, "0")
                End If
                
                AVP# = TaxeRec!vk / 100#
                h$ = PznString(TaxeRec!pzn) + vbTab + TaxeRec!Name + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit
                h$ = h$ + vbTab + TaxeRec!HerstellerKB + vbTab + POSStr$
                h$ = h$ + vbTab + Format(AVP#, "0.00") + vbTab + Format(NextMulti%, "0") + vbTab + Format(AVP# * CDbl(NextMulti%), "0.00")
                flxOrg.AddItem h$
 '           End If
        End If
      End If
    End If
    
    found& = found& + 1
    j& = j& + 1
    
    If ((j& Mod 100) = 0) Then
        DoEvents
        If (AbbruchSuche%) Then
            flxOrg.Rows = 1
            Exit Do
        End If
    End If
Loop

h$ = ""
'7 Multi
'8 Wert

With flxOrg
    If (.Rows > 1) Then
        'sortieren nach PZN
        .row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = 0
        .Sort = 5
        
        l& = 1
        Do
            If (l& >= .Rows) Then Exit Do
            
            If (.TextMatrix(l&, 0) = h$) Then
                .TextMatrix(l& - 1, 7) = Format(xVal(.TextMatrix(l& - 1, 7)) + xVal(.TextMatrix(l&, 7)), "0")
                .TextMatrix(l& - 1, 8) = Format(xVal(.TextMatrix(l& - 1, 8)) + xVal(.TextMatrix(l&, 8)), "0.00")
                
                .RemoveItem l&
            Else
                h$ = .TextMatrix(l&, 0)
                l& = l& + 1
            End If
        Loop

        l& = 1
        Do
            If (l& >= .Rows) Then Exit Do
            
            LoeschFlag% = False
            If (Anzahl% > 0) Then
                If (xVal(.TextMatrix(l&, 7)) < Anzahl%) Then
                    If (Wert# = 0) Or (xVal(.TextMatrix(l&, 8)) < Wert#) Then LoeschFlag% = True
                End If
            ElseIf (Wert# > 0#) Then
                If (xVal(.TextMatrix(l&, 8)) < Wert#) Then LoeschFlag% = True
            End If
            
            If (LoeschFlag%) Then
                If (.Rows <= 2) Then
                    For i% = 0 To (.Cols - 1)
                        .TextMatrix(l&, i%) = " "
                    Next i%
                    .TextMatrix(l&, 1) = "keine passenden Daten vorhanden !"
                    Exit Do
                Else
                    .RemoveItem l&
                End If
            Else
                l& = l& + 1
            End If
        Loop
    End If
    
    If (.Rows <= 1) Then .AddItem vbTab + "keine passenden Daten vorhanden !"
    
    Call SortFlex
    
    cmdF3.Enabled = (Val(.TextMatrix(1, 0)) > 0)
    cmdF6.Enabled = cmdF3.Enabled
    If (cmdF6.Enabled) Then .SetFocus
    
End With

MousePointer = vbNormal
InSuche% = False

Call StopAnimation(Me)

Call DefErrPop
End Sub

Sub LagerOriginal()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LagerOriginal")
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
Dim satz As Long
Dim SQLStr As String
Dim OrgPZN As Long
Dim OrgAvp As Double
Dim aktavp As Double
Dim POSStr As String
Dim MeinAVP As Double
Dim ok As Boolean
Dim rr As Integer
Dim AVP#

Dim j&, l&, h$, LoeschFlag%

Call StartAnimation(Me, "Alt.Importe werden gesucht ...")

InSuche% = True
AbbruchSuche% = False
MousePointer = vbHourglass

flxOrg.Rows = 1
flxOrg.Redraw = False

ass.GetRecord (1)
For satz = 1 To ass.erstmax
  ass.GetRecord (satz + 1)
  ok = False
  SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ass.pzn
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
  If Not TaxeRec.EOF Then
    If InStr("0,2,3,4", CStr(TaxeRec!Abgabebest)) = 0 Then ok = True    'RezPfl
    MeinAVP = TaxeRec!vk / 100#
  End If
  If ok Then
    If (IstOriginal(Val(ass.pzn), OrgPZN)) Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Str$(OrgPZN)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If Not TaxeRec.EOF Then
'            If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then    'GS: muss kein Original sein, um als solches zu gelten
                POSStr$ = Format(ass.poslag, "0")
                
                AVP# = TaxeRec!vk / 100#
                h$ = PznString(TaxeRec!pzn) + vbTab + TaxeRec!Name + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit
                h$ = h$ + vbTab + TaxeRec!HerstellerKB + vbTab + POSStr$
                h$ = h$ + vbTab + Format(AVP#, "0.00") + vbTab + vbTab
                flxOrg.AddItem h$
 '           End If
        End If
'      SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(OrgPZN)
'      Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'      If Not TaxeRec.NoMatch Then
'        OrgAvp = TaxeRec!VK / 100#
'      End If
'
'      SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + CStr(OrgPZN)
'      Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'      If Not TaxeRec.NoMatch Then
'        TaxeRec.MoveFirst
'        Do
'          If TaxeRec!pzn <> OrgPZN And TaxeRec!pzn <> Val(ass.pzn) Then
'            aktavp = TaxeRec!VK / 100#
'            If aktavp <= (OrgAvp - 15) Or aktavp <= (OrgAvp * 0.85) Then
'
'              'das ist ein Import
'
'              If Not ok Then
'                POSStr$ = Format(ass.poslag, "0")
'                FabsErrf% = ast.IndexSearch(0, ass.pzn, FabsRecno&)
'                If (FabsErrf% = 0) Then
'                  ast.GetRecord (FabsRecno& + 1)
'                End If
'                h$ = Format(OrgPZN, "0000000") + vbTab + ast.kurz + vbTab + ast.meng + vbTab + ast.meh
'                h$ = h$ + vbTab + ast.herst + vbTab + POSStr$ + vbTab + Format(MeinAVP, "0.00") + vbTab + "o"
'                flxOrg.AddItem h$
'                flxOrg.row = flxOrg.Rows - 1
'                For rr = 0 To flxOrg.Cols - 1
'                  flxOrg.col = rr
'                  flxOrg.CellBackColor = vbRed
'                Next rr
'                ok = True
'              End If
'              POSStr$ = ""
'              FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
'              If (FabsErrf% = 0) Then
'                ass.GetRecord (FabsRecno& + 1)
'                POSStr$ = Format(ass.poslag, "0")
'              End If
'              h$ = Format(TaxeRec!pzn, "0000000") + vbTab + TaxeRec!Name + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit
'              h$ = h$ + vbTab + TaxeRec!HerstellerKB + vbTab + POSStr$
'              h$ = h$ + vbTab + Format(aktavp#, "0.00")
'              flxOrg.AddItem h$
'              If POSStr$ > "" Then
'                flxOrg.row = flxOrg.Rows - 1
'                For rr = 0 To flxOrg.Cols - 1
'                  flxOrg.col = rr
'                  flxOrg.CellFontBold = True
'                Next rr
'              End If
'            End If
'          End If
'          TaxeRec.MoveNext
'        Loop Until TaxeRec.EOF
'      End If  'If Not TaxeRec.NoMatch Then
    End If  'If IstOriginal(Val(ass.pzn), OrgPZN) Then
  End If    'if ok then

  j& = j& + 1
  
  If ((j& Mod 100) = 0) Then
      DoEvents
      If (AbbruchSuche%) Then
          flxOrg.Rows = 1
          Exit For
      End If
  End If

Next satz

h$ = ""
'7 Multi
'8 Wert

With flxOrg
    If (.Rows <= 1) Then .AddItem vbTab + "keine passenden Daten vorhanden !"
    
    Call SortFlex
    
    cmdF3.Enabled = False
    cmdF6.Enabled = (Val(.TextMatrix(1, 0)) > 0)
    If (cmdF6.Enabled) Then .SetFocus
End With

MousePointer = vbNormal
InSuche% = False

Call StopAnimation(Me)

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

If (KeyCode = vbKeyF3) Then
    cmdF3.Value = True
ElseIf (KeyCode = vbKeyF6) Then
    cmdF6.Value = True
End If

Call DefErrPop
End Sub

Sub ImpAlternativAusdruck(Optional ManuellDruck% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImpAlternativAusdruck")
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
Dim i%, j%, k%, m%, ret%, rInd%, y%, sp%(5), anz%, ind%, AktLief%, Erst%, AnzRetourArtikel%, Handle%
Dim ZentrierX%, x%, OrgAttrib%, ImpAusdruckImporte%, ZeilenHöhe%
Dim gesBreite&, BlockHe&
Dim RetourWert#, AVP#
Dim tx$, h$
Dim pzn$, SQLStr$

'ImpAusdruckImporte% = 0
'
'ret% = MsgBox("Ausdruck inklusive Import-Alternativen ?", vbYesNoCancel Or vbQuestion, "Import-Alternativen")
'If (ret% = vbCancel) Then Call DefErrPop: Exit Sub
'
'If (ret% = vbYes) Then
'    h$ = InputBox("Anzahl Import-Alternativen pro Original: ", "Ausdruck Import-Alternativen", "5")
'    If (Val(h$) <= 0) Then
'        ret% = vbCancel
'    Else
'        ImpAusdruckImporte% = Val(h$)
'    End If
'End If
'If (ret% = vbCancel) Then Call DefErrPop: Exit Sub


Call StartAnimation(Me, "Ausdruck wird erstellt ...")

If (ImpAlternativModus% = 0) Then
    AnzDruckSpalten% = 10
Else
    AnzDruckSpalten% = 8
End If
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .Titel = " "
    .TypStr = String$(1, "X")
    .Ausrichtung = "L"
    .Attrib = 1
End With
With DruckSpalte(1)
    .Titel = "P Z N"
    .TypStr = String$(8, "9")
    .Ausrichtung = "L"
End With
With DruckSpalte(2)
    .Titel = "A R T I K E L"
    .TypStr = String$(24, "X")  '28
    .Ausrichtung = "L"
End With
With DruckSpalte(3)
    .Titel = ""
    .TypStr = String$(5, "X")
    .Ausrichtung = "R"
End With
With DruckSpalte(4)
    .Titel = ""
    .TypStr = String$(3, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(5)
    .Titel = "Herst"
    .TypStr = String$(7, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(6)
    .Titel = "POS"
    .TypStr = String$(3, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(7)
    .Titel = "A V P"
    .TypStr = "99999.99"
    .Ausrichtung = "R"
End With
If (ImpAlternativModus% = 0) Then
    With DruckSpalte(8)
        .Titel = "Anz"
        .TypStr = String$(5, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(9)
        .Titel = "Wert"
        .TypStr = "999999.99"
        .Ausrichtung = "R"
    End With
End If


Call InitDruckZeile(True)

DruckSeite% = 0
Call ImpAlternativDruckKopf
ZeilenHöhe% = Printer.TextHeight("A")

flxImporte.Redraw = False
For i% = 1 To (flxOrg.Rows - 1)
    pzn$ = flxOrg.TextMatrix(i%, 0)
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If Not TaxeRec.EOF Then
        AVP# = TaxeRec!vk / 100#
        h$ = vbTab + PznString(TaxeRec!pzn)
        h$ = h$ + vbTab + TaxeRec!Name + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit
        h$ = h$ + vbTab + TaxeRec!HerstellerKB + vbTab + flxOrg.TextMatrix(i%, 5)
        h$ = h$ + vbTab + Format(AVP#, "0.00")
        h$ = h$ + vbTab + flxOrg.TextMatrix(i%, 7) + vbTab + flxOrg.TextMatrix(i%, 8) + vbTab
    
        BlockHe& = ImpAusdruckImporte% * Printer.TextHeight("H")
        If ((Printer.CurrentY + BlockHe&) > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call ImpAlternativDruckKopf
        End If
        
        
        Call DruckZeile(h$)
        
        Printer.FontSize = DruckFontSize% - 1
        Printer.FontItalic = True
        Call GetImporte(pzn$)
        With flxImporte
            j% = 0
            For m% = .FixedRows To (.Rows - 1)
                .row = m%
                .col = 0
'                If (.CellFontBold Or (j% < ImpAusdruckImporte%)) Then
                If (.CellBackColor <> vbGrayText) Then
                    Printer.FontBold = .CellFontBold
                    h$ = "I" + vbTab
                    For k% = 0 To 6
                        h$ = h$ + .TextMatrix(m%, k%) + vbTab
                    Next k%
'                    h$ = h$ + .TextMatrix(m%, 6) + vbTab + .TextMatrix(m%, 8) + vbTab + .TextMatrix(m%, 5) + vbTab
                    h$ = h$ + vbTab + vbTab
                    Call DruckZeile(h$)
                    Printer.FontBold = False
                    j% = j% + 1
                End If
            Next m%
        End With
        Printer.FontSize = DruckFontSize%
        Printer.FontItalic = False

        Printer.CurrentY = Printer.CurrentY + 150
    End If
    
    If (Printer.CurrentY > Printer.ScaleHeight - 1000 - ZeilenHöhe%) Then
        Call DruckFuss
        Call ImpAlternativDruckKopf
    End If
Next i%
    
Call DruckFuss(False)
Printer.EndDoc

Call StopAnimation(Me)

flxImporte.Redraw = True
flxOrg.SetFocus
                
Call DefErrPop
End Sub

Sub GetImporte(pzn$, Optional MitRedraw% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetImporte")
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
Dim OrgAvp#, AVP#
Dim s$, SQLStr$, kz$, POSStr$

With flxImporte
    If (MitRedraw%) Then
        .Redraw = False
    End If
    .Rows = .FixedRows
    
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If TaxeRec.EOF Then
        Call DefErrPop: Exit Sub
    End If
    
    OrgAvp# = TaxeRec!vk / 100#

    SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Str$(TaxeRec!OriginalPZN)
    
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    
    While Not (TaxeRec.EOF)
        AVP# = TaxeRec!vk / 100#
    
        If (TaxeRec!pzn <> Val(pzn$)) Then
'            If (avp# < (OrgAvp# - 1)) And (avp# < (OrgAvp# * 0.9)) Then    'GS
            If (AVP# < (OrgAvp# - 15)) Or (AVP# < (OrgAvp# * 0.85)) Then
                kz$ = ""
            Else
                kz$ = "-"
            End If
            
'            If kz$ = "" Then    ' GS nur entsprechend günstige Importe gelten als Importe
            POSStr$ = ""
            If (ArtikelDbOk) Then
                SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
            '                    SQLStr = SQLStr + " AND LagerKz<>0"
                FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
            Else
                FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
                If (FabsErrf% = 0) Then
                    ass.GetRecord (FabsRecno& + 1)
                End If
            End If
            If (FabsErrf% = 0) Then
                POSStr$ = Format(ass.poslag, "0")
            End If
                
            s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab
            s$ = s$ + TaxeRec!HerstellerKB + vbTab + POSStr$ + vbTab
            s$ = s$ + Format(AVP#, "0.00") + vbTab
            s$ = s$ + Format(AVP#, "000000.00") + Trim(TaxeRec!Name) + TaxeRec!menge + TaxeRec!HerstellerKB + vbTab
            .AddItem s$
                
            If (POSStr$ <> "") Then
                .FillStyle = flexFillRepeat
                .row = .Rows - 1
                .col = 0
                .RowSel = .row
                .ColSel = .Cols - 1
                .CellFontBold = True
                .FillStyle = flexFillSingle
            End If
            
            If (kz$ <> "") Then
                .FillStyle = flexFillRepeat
                .row = .Rows - 1
                .col = 0
                .RowSel = .row
                .ColSel = .Cols - 1
                .CellBackColor = vbGrayText
                .FillStyle = flexFillSingle
            End If
            
        End If
    
        TaxeRec.MoveNext
    Wend

    .row = 1
    .col = 7
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    .col = 1
    .ColSel = .Cols - 1
    
    For i% = .FixedRows To (.Rows - 1)
        .TextMatrix(i%, 7) = ""
    Next i%
    
    .col = 0
    .ColSel = .Cols - 1
    If (MitRedraw%) Then
        .Redraw = True
    End If
End With

Call DefErrPop
End Sub

Sub ImpAlternativDruckKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImpAlternativDruckKopf")
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
Dim i%, x%, y%, AbAnzPackungen%
Dim gesBreite&
Dim AbGesamtWert#
Dim header$, KopfZeile$, Typ$, h$

KopfZeile$ = "Import-Kontrolle"
If (ImpAlternativModus% = 0) Then
    header$ = "Abgegebene Original-Artikel"
    Typ$ = "von " + ImpAlternativPara$(0) + " bis " + Format(Now, "DDMMYY")

    AbAnzPackungen% = Val(ImpAlternativPara$(1))
    If (AbAnzPackungen% > 0) Then
        Typ$ = Typ$ + ", ab" + Str$(AbAnzPackungen%) + " Packungen"
    End If
    
    AbGesamtWert# = Val(ImpAlternativPara$(2))
    If (AbGesamtWert# > 0) Then
        Typ$ = Typ$ + ", ab" + Str$(AbGesamtWert#) + " EUR Wert"
    End If
Else
    header$ = "Import-Alternativen"
    Typ$ = ""
End If
        
Call DruckKopf(header$, Typ$, KopfZeile$, 14)
    
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

Private Sub flxOrg_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOrg_GotFocus")
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

With flxOrg
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxOrg_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOrg_LostFocus")
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

With flxOrg
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub flxOrg_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOrg_RowColChange")
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
Dim pzn$

With flxOrg
    If (.Redraw) Then
        pzn$ = .TextMatrix(.row, 0)
        If (Val(pzn$) > 0) Then
            Call GetImporte(pzn$)
        End If
    End If
End With

Call DefErrPop
End Sub

Sub InitAnimation()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitAnimation")
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
With lblAnimation
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = TextWidth("Parameter werden eingelesen ...") + 300
    .Height = TextHeight("Äg") + 150
End With

With aniAnimation
    .Left = lblAnimation.Left + (lblAnimation.Width - .Width) / 2
    .Top = lblAnimation.Top + lblAnimation.Height + 90
End With

With picAnimationBack
    .Width = lblAnimation.Width + 2 * wpara.LinksX
    .Height = aniAnimation.Top + aniAnimation.Height + 90
End With

Call DefErrPop
End Sub


