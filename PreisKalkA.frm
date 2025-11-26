VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPreisKalkA 
   Caption         =   "Preis-Korrektur"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdF8 
      Caption         =   "&Preiskalkulation (F8)"
      Height          =   450
      Left            =   2160
      TabIndex        =   3
      Top             =   4440
      Width           =   1200
   End
   Begin VB.CommandButton cmdF7 
      Caption         =   "&Aufschlagtabelle (F7)"
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid FlxPreisKorrektur 
      Height          =   2775
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4895
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   5880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1200
   End
End
Attribute VB_Name = "frmPreisKalkA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
        
'Private Type WkZwStruct
'    pzn As String * 7
'    datum As String * 2
'    lief As String * 9
'    menge As Integer
'    KontrNr As String * 11
'    DruKz As String * 1
'    Meng2 As String * 5
'    Name As String * 2
'    IdDatum As String * 2
'    Preis As Double
'    Pruef As String * 18
'    rest As String * 61
'End Type
'Dim WkZw As WkZwStruct

Dim PkErr%
Dim PkErrTxt$

Dim PkAep#
Dim PkKp#
Dim PkAvp#
Dim PkMw!
Dim PkWg$
Dim PkLc$
Dim PkKas$
Dim PkKasz$
Dim PkRez$
Dim PkAg$
Dim PkAvpMin#
Dim PkAvpMax#

Dim argv$()

Private Const DefErrModul = "PREISKALKA.FRM"

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
Dim i%, OrgRedraw%, OldRow%, rInd%, row%, wg%, AlleKontrNr%, AlleBesorgerAvp%
Dim iAvp#

With FlxPreisKorrektur
    AlleKontrNr% = True
    For i% = 1 To (.Rows - 1)
        wg% = xVal(.TextMatrix(i%, 8))
        If (wg% = 3) Or ((wg% >= 30) And (wg% <= 39)) Then
            If (Trim(.TextMatrix(i%, 7)) = "") Then
                AlleKontrNr% = False
                Exit For
            End If
        End If
    Next i%
    If (AlleKontrNr% = False) Then
        Call MsgBox("Achtung: Noch nicht alle Kontroll-Nummern eingegeben!", vbExclamation)
        Call DefErrPop: Exit Sub
    End If
    
    AlleBesorgerAvp% = True
    For i% = 1 To (.Rows - 1)
        row% = xVal(.TextMatrix(i%, 10))
        frmAction.flxarbeit(0).row = row%
        
        rInd% = SucheFlexZeile(True)
        If (rInd% > 0) Then
            If (ww.AbholNr > 0) Then
                iAvp# = xVal(.TextMatrix(i%, 5))
                If (iAvp# <= 0#) Then
                    AlleBesorgerAvp% = False
                    Exit For
                End If
            End If
        End If
    Next i%
    If (AlleBesorgerAvp% = False) Then
        Call MsgBox("Achtung: Noch nicht alle Besorger-AVPs eingegeben!", vbExclamation)
        Call DefErrPop: Exit Sub
    End If
    
    With frmAction.flxarbeit(0)
        OrgRedraw% = .redraw
        .redraw = False
        OldRow% = .row
    End With
    
    ww.SatzLock (1)
    For i% = 1 To (.Rows - 1)
        row% = xVal(.TextMatrix(i%, 10))
        frmAction.flxarbeit(0).row = row%
        
        rInd% = SucheFlexZeile(True)
        If (rInd% > 0) Then
            wg% = xVal(.TextMatrix(i%, 8))
            If (wg% = 3) Or ((wg% >= 30) And (wg% <= 39)) Then
                Call WriteWZ(i%)
            End If
            
            Call ActProgram.SpeicherKalkPreise(rInd%)
            
'            If (ww.WuStatus = 1) Then
'                ww.WuStatus = 2
'            End If
'            ww.PutRecord (rInd% + 1)
        End If
    Next i%
    ww.SatzUnLock (1)
    
    With frmAction.flxarbeit(0)
        .redraw = OrgRedraw%
        .row = OldRow%
    End With
    
End With
Unload Me

Call DefErrPop
End Sub

Private Sub cmdF7_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF7_Click")
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

Call AufschlagAuswahl

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
Dim PkInd%
Dim aep#, AVP#, Wert#
Dim errtxt$

PkInd% = SelectPk%
If (PkInd% > 0) Then
    With FlxPreisKorrektur
        aep# = xVal(.TextMatrix(.row, 4))
        Call PkBelegen(aep#)
        If PKkalk%(PkInd%, errtxt$, Wert#) = 0 Then
            AVP# = Wert#
            .TextMatrix(.row, 5) = Format(AVP#, "0.00")
            .col = 5
        End If
    End With
    Call PreisKorrekturSpeichern
End If

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
Dim ind%

If (KeyCode = vbKeyF7) Then
    cmdF7.Value = True
ElseIf (KeyCode = vbKeyF8) Then
    cmdF8.Value = True
ElseIf (ActiveControl.Name = FlxPreisKorrektur.Name) And (KeyCode = vbKeyReturn) Then
    Call EditSatz
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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%
Dim h$, h2$

Call wpara.InitFont(Me)

PreisKalkAepChange% = False

With FlxPreisKorrektur
    .Cols = 12
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|<Name|>Menge|^Meh|>AEP|>AVP|>PkNr|>KontrollNr|>Wg|>A||"
    .Rows = 1
    
        
    Font.Bold = True
    .ColWidth(0) = TextWidth("99999999")
    .ColWidth(1) = TextWidth("Xxxxxx Xxxxxx Xxxxxx Xxxxxx Xxxxx")
    .ColWidth(2) = TextWidth("XXXXXXX")
    .ColWidth(3) = TextWidth("XXXX")
    .ColWidth(4) = TextWidth("9999999.99")
    .ColWidth(5) = TextWidth("9999999.99")
    .ColWidth(6) = TextWidth("99999")
    .ColWidth(7) = TextWidth(String(13, "9"))
    .ColWidth(8) = TextWidth(String(4, "9"))
    .ColWidth(9) = TextWidth(String(3, "9"))
    .ColWidth(10) = 0
    .ColWidth(11) = wpara.FrmScrollHeight
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    
    Call PreisKalkBefuellen
'    .Rows = 1
    If (.Rows = 1) Then .Rows = 2

    .row = 1
    .col = 1
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = 3
    .Sort = 5

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = 11 * .RowHeight(0) + 90

    .SelectionMode = flexSelectionFree
    
    .row = 1
    .col = 4
'    .ColSel = .Cols - 1
End With
    

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With cmdF7
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = FlxPreisKorrektur.Left
    .Top = FlxPreisKorrektur.Top + FlxPreisKorrektur.Height + 150 * wpara.BildFaktor
End With

With cmdF8
    .Width = cmdF7.Width
    .Height = cmdF7.Height
    .Left = cmdF7.Left + cmdF7.Width + 150
    .Top = cmdF7.Top
End With

Me.Width = FlxPreisKorrektur.Left + FlxPreisKorrektur.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = cmdF7.Top
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = FlxPreisKorrektur.Left + FlxPreisKorrektur.Width - .Width
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


Call DefErrPop
End Sub
    
Private Sub PreisKalkBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisKalkBefuellen")
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
Dim i%, AltRow%, rInd%
    
With frmAction.flxarbeit(0)
    .redraw = False
    AltRow% = .row
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = "$") Then
            .row = i%
            rInd% = SucheFlexZeile(False)
            If (rInd% > 0) Then
                Call PreisKalkZeile(rInd%, i%)
            End If
        End If
    Next i%
    .row = AltRow%
    .redraw = True
End With


Call DefErrPop
End Sub

Private Sub PreisKalkZeile(rInd%, pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisKalkZeile")
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
Dim j%, l%, PkInd%, PosLag%, wg%, Aufschlag%
Dim xAep#, xavp#
Dim text$, pzn$, h$, h2$, tax$, KzAuf$
Dim ArtName$, ArtMenge$, ArtMeh$
Dim PreisFaktor

pzn$ = ww.pzn
        
'If xAep# <= 0# Then xAep# = 0.01
'If xavp# <= 0# Then xavp# = 0.01

PkInd% = 0
PreisFaktor = 1: wg% = 0: tax$ = "  ": KzAuf$ = " "
Aufschlag% = 0

text$ = ww.txt
If (pzn$ = "9999999") Then
    h$ = RTrim$(text$)
    Call OemToChar(h$, h$)
    ArtName$ = h$
    ArtMenge$ = ""
    ArtMeh$ = ""
Else
    h$ = RTrim(Left$(text$, 33))
    Call OemToChar(h$, h$)
    l% = Len(h$)
    For j% = l% To 2 Step -1
        h2$ = Mid$(h$, j%, 1)
        If (h2$ = " ") Then Exit For
        If (InStr(".xX", h2$) > 0) Then
            h2$ = Mid$(h$, j% - 1, 1)
            If (InStr("0123456789", h2$) <= 0) Then
                Exit For
            End If
        End If
    Next j%
    If (j% > 2) Then
        ArtName$ = Left$(h$, j%)
        ArtMenge$ = Mid$(h$, j% + 1)
    Else
        ArtName$ = h$
        ArtMenge$ = ""
    End If
    ArtMeh$ = Mid$(text$, 34)
End If

xAep# = 0#
xavp# = ww.WuAVP
If (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        xAep# = ww.WuAEP    ' ast.aep
        '2.96 AVP in Stammdaten meist aktueller als in WÜ-Satz
        If (ast.AVP <> 0) Then xavp# = ast.AVP
'        mw$ = ast.mw
        
        wg% = Val(ast.wg1): tax$ = ast.wg
        KzAuf$ = ast.ka: If (Asc(KzAuf$) = 0) Then KzAuf$ = " "
        Aufschlag% = Val(ast.ka)
'        If (wg% = 3) And (Val(ast.herst) <> 0) And (Val(ast.meng) <> 0) Then
'            PreisFaktor = Val(ast.herst) / Val(ast.meng)
'        End If
    End If
End If
xAep# = xAep# * PreisFaktor
xavp# = xavp# * PreisFaktor

PosLag% = 0
PkInd% = 0
FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Call ass.GetRecord(FabsRecno& + 1)
    PosLag% = ass.PosLag
    PkInd% = ass.pk
    If (PkInd% = 8224) Then PkInd% = 0
End If
'If (ww.auto = "v") And (AufschlagsTabelle(MAX_AUFSCHLAEGE - 1).PreisBasis <> 0) Then PkInd% = MAX_AUFSCHLAEGE



h$ = pzn$ + vbTab + ArtName$ + vbTab + ArtMenge$ + vbTab + ArtMeh$ + vbTab
h$ = h$ + Format(xAep#, "0.00") + vbTab + Format(xavp#, "0.00") + vbTab
If (PkInd% > 0) Then
    h$ = h$ + Str$(PkInd%)
End If
h$ = h$ + vbTab
h$ = h$ + vbTab
h$ = h$ + tax$ + vbTab + KzAuf$ + vbTab + Str$(pos%)

FlxPreisKorrektur.AddItem h$

Call DefErrPop
End Sub

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iKalk%, mw%, aufschl%, aCol%, iKalkModus%
Dim PkInd%
Dim KalkAvp#, aep#, AVP#, Wert#
Dim KalkText$, Col1$, errtxt$
Dim h$, h2$
            
EditModus% = 4
            
EditCol% = FlxPreisKorrektur.col
If (EditCol% = 4) Or (EditCol% = 5) Or (EditCol% = 7) Then
            
    With FlxPreisKorrektur
        aRow% = .row
        .row = 0
        .CellFontBold = True
        .row = aRow%
    End With
            
    Load frmEdit
    
    If (EditCol% = 7) Then
        EditModus% = 1
    End If
    
    With frmEdit
        .Left = FlxPreisKorrektur.Left + FlxPreisKorrektur.ColPos(EditCol%) + 45
        .Left = .Left + Left + wpara.FrmBorderHeight
        .Top = FlxPreisKorrektur.Top + (FlxPreisKorrektur.row - FlxPreisKorrektur.TopRow + 1) * FlxPreisKorrektur.RowHeight(0)
        .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = FlxPreisKorrektur.ColWidth(EditCol%)
        .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
    End With
    With frmEdit.txtEdit
        .Width = frmEdit.ScaleWidth
'            .Height = frmEdit.ScaleHeight
        .Left = 0
        .Top = 0
        h2$ = FlxPreisKorrektur.TextMatrix(FlxPreisKorrektur.row, EditCol%)
        .text = h2$
        .BackColor = vbWhite
        .Visible = True
    End With
   
    frmEdit.Show 1
            
    With FlxPreisKorrektur
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
            
        If (EditErg%) Then
            If (EditCol% = 4) Then
                aep# = xVal(EditTxt$)
                .TextMatrix(.row, EditCol%) = Format(aep#, "0.00")
                PkInd% = Val(.TextMatrix(.row, 6))
                If (PkInd% > 0) Then
                    Call PkBelegen(aep#)
                    If PKkalk%(PkInd%, errtxt$, Wert#) = 0 Then
                        AVP# = Wert#
                        .TextMatrix(.row, 5) = Format(AVP#, "0.00")
                        .col = 5
                    End If
                End If
                Call PreisKorrekturSpeichern
            ElseIf (EditCol% = 5) Then
                AVP# = xVal(EditTxt$)
                .TextMatrix(.row, EditCol%) = Format(AVP#, "0.00")
                Call PreisKorrekturSpeichern
            Else
                .TextMatrix(.row, EditCol%) = EditTxt$
            End If
        End If
    End With
End If

Call DefErrPop
End Sub


Function PKkalk%(pknr%, errtxt$, Wert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PKkalk%")
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
Dim fPk%, z%, t%
Dim Zeile$(20), zStr$
Dim RechenZeile%

PKkalk% = True 'Fehler!

fPk% = FreeFile
'On Error Resume Next
Open "PKSCHEMA.TXT" For Input As #fPk%
If Err = 0 Then
    RechenZeile% = 1
    While Not EOF(fPk%)
        Line Input #fPk%, zStr$
        Call SeparateKomma(zStr$)
        If (pknr% = Val(argv$(1))) Then
            z% = Val(argv$(2))
            If (z% >= 1) And (z% <= 20) Then
                t% = Val(argv$(3))
                argv$(4) = Mid$(argv$(4), 2, Len(argv$(4)) - 2)
                If (t% = 1) Then
                    RechenZeile% = 0
                    Zeile$(z% * 2 - 1) = argv$(4)
                    If PKwert%(argv$(4), errtxt$, Wert#) Then
                      errtxt$ = "Fehler in Bedingung!"
                      PKkalk% = True
                      Close #fPk%
                      Call DefErrPop: Exit Function
                  Else
                    If (Wert# <> 0#) Then
                      RechenZeile% = z%
                    End If
                  End If
                Else
                    If (RechenZeile% = z%) Then
                        If PKwert%(argv$(4), errtxt$, Wert#) Then
                            errtxt$ = "Fehler in Bedingung!"
                            PKkalk% = True
                            Close #fPk%
                            Call DefErrPop: Exit Function
                        Else
                            PKkalk% = 0
                            Close #fPk%
                            Call DefErrPop: Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Wend
    Close #fPk%
End If

errtxt$ = "Kalkulationsschema nicht gefunden!"
'On Error GoTo 0

Call DefErrPop
End Function

Sub SeparateKomma(z As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeparateKomma")
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
Dim i As Integer
Dim argc As Integer
Dim s As String
Dim oldp%, p%, x1%, x2%

ReDim argv(1) As String
s = z
argc = 0

Do
    oldp% = 0
    p% = 1
    While (p% <> oldp%)
        x2% = 0
        x1% = InStr(s, Chr$(34))
        If (x1% > 0) Then
            x2% = InStr(x1% + 1, s, Chr$(34))
        End If
        i = InStr(p%, s, ",")
        oldp% = p%
        If (i > x1%) And (i < x2%) Then
            p% = x2% + 1
        End If
    Wend
    
    If (i > 0) Then
        argc = argc + 1
        ReDim Preserve argv(argc) As String
        argv(argc) = LTrim$(RTrim$(Left$(s, i - 1)))
        s = Mid$(s, i + 1)
    Else
        argc = argc + 1
        ReDim Preserve argv(argc) As String
        argv(argc) = LTrim$(RTrim$(s))
        s = ""
    End If
Loop While (s <> "")

Call DefErrPop
End Sub


Function PKwert%(EingabeZeile$, errtxt$, Wert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PKwert%")
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
Dim b$

PkErr% = 0
Wert# = 0#
b$ = UCase$(RTrim$(EingabeZeile$))
If (Len(b$) > 0) Then
    Call PKkonvert(b$)
    Wert# = Val(PKparse$(b$))
Else
    Wert# = -1
End If
PKwert% = PkErr%
errtxt$ = PkErrTxt$

Call DefErrPop
End Function

Sub PKkonvert(b$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PKkonvert")
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
Dim X%, p1%, p2%, break%
Dim c$, x2$

Do
    X% = InStr(b$, " UND ")
    If (X% > 0) Then
        b$ = Left$(b$, X% - 1) + " & " + Mid$(b$, X% + 5)
    End If
Loop While (X% > 0)

Do
    X% = InStr(b$, " ODER ")
    If (X% > 0) Then
        b$ = Left$(b$, X% - 1) + " | " + Mid$(b$, X% + 6)
    End If
Loop While (X% > 0)

Do
    X% = InStr(b$, " IN ")
    If (X% > 0) Then
        b$ = Left$(b$, X% - 1) + " ~ " + Mid$(b$, X% + 3)
    End If
Loop While (X% > 0)

'alle +x% oder -x% durch *(1+x/100) oder *(1-x/100) ersetzen
p1% = 1
While (p1% > 0)
    p1% = InStr(b$, "%")
    If (p1% > 0) Then
        break% = 0
        p2% = p1%
        While (p1% > 1) And Not break%
            p1% = p1% - 1
            c$ = Mid$(b$, p1%, 1)
            If (c$ = "+") Or (c$ = "-") Then
                break% = True
            End If
        Wend
        If (p1% > 1) Then
            x2$ = Mid$(b$, p1% + 1, p2% - p1% - 1)
            b$ = Left$(b$, p1% - 1) + "*" + "(1" + c$ + x2$ + "/100)" + Mid$(b$, p2% + 1)
        End If
    End If
Wend

Call DefErrPop
End Sub

Function PKparse$(Zeile$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PKparse$")
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
Dim i%, p%, klammer%, KlammerPaar%, minop%, op%, op2%, opPos%, opLen%, numerisch%
Dim w#
Dim erg$, b$, c$, fuName$, wert1$, wert2$

erg$ = ""
PKparse$ = erg$

b$ = LTrim$(RTrim$(Zeile$))

If (b$ = "") Then Call DefErrPop: Exit Function

'Klammern entfernen
If (Left$(b$, 1) = "(") And (Right$(b$, 1) = ")") Then
    p% = 1: klammer% = 0
    Do
        c$ = Mid$(b$, p%, 1)
        If (c$ = "(") Then
            klammer% = klammer% + 1
        ElseIf (c$ = ")") Then
            klammer% = klammer% - 1
            If (klammer% = 0) Then
                If (p% = Len(b$)) Then b$ = Mid$(b$, 2, Len(b$) - 2)
                Exit Do
            End If
        End If
        p% = p% + 1
    Loop While (p% <= Len(b$))
End If

'niederwertigsten Operator suchen
p% = 1
klammer% = 0
minop% = 99
Do
    c$ = Mid$(b$, p%, 1)
    If (c$ = "(") Then
        klammer% = klammer% + 1
    ElseIf (c$ = ")") Then
        klammer% = klammer% - 1
        KlammerPaar% = KlammerPaar% + 1
    Else
        op% = InStr("|&-+/*~", c$)
        If (klammer% = 0) Then
            If (op% > 0) Then
                If (op% < minop%) Or (minop% > 9) Then
                    minop% = op%: opPos% = p%: opLen% = 1
                End If
            Else
                op% = InStr("<=>", c$)
                If (op% > 0) And (minop% = 99) Then
                    opPos% = p%: opLen% = 1: op% = op% * 10
                    op2% = InStr("<=>", Mid$(b$, p% + 1, 1))
                    If (op2% > 0) Then
                        op% = op% + op2%: opLen% = 2: p% = p% + 1
                    End If
                    minop% = op%
                End If
            End If
        End If
    End If
    p% = p% + 1
Loop While (p% <= Len(b$))

If (minop% < 99) Then
    wert1$ = PKparse$(Left$(b$, opPos% - 1))
    If PkErr% Then Call DefErrPop: Exit Function
    wert2$ = PKparse$(Mid$(b$, opPos% + opLen%))
    If PkErr% Then Call DefErrPop: Exit Function
Else
'Funktion?
    If (KlammerPaar% > 0) Then
        p% = InStr(b$, "(")
        If (p% = 0) Then
            PkErrTxt$ = b$ + Chr$(0) + " -> Syntaxfehler!"
            PkErr% = True
        End If
        fuName$ = Left$(b$, p% - 1)
        b$ = Mid$(b$, p%)
        Select Case fuName$
            Case "RUND"
                wert1$ = PKfuparam$(b$, 1)
                If PkErr% Then Call DefErrPop: Exit Function
                wert2$ = PKfuparam$(b$, 2)
                If PkErr% Then Call DefErrPop: Exit Function
                w# = Val(wert2$)
                If (w# > 0#) Then
                    w# = 1 / w#
                    PKparse$ = Str$(Int(Val(wert1$) * w# + 0.5) / w#)
                    Call DefErrPop: Exit Function
                End If
                PkErrTxt$ = fuName$ + ":" + b$ + Chr$(0) + " -> Parameter 2 ist 0!"
            Case Else
                PkErrTxt$ = fuName$ + Chr$(0) + " -> Funktion unbekannt!"
                PkErr% = True
                Call DefErrPop: Exit Function
        End Select
    End If
End If


Select Case minop%
    Case 1
        erg$ = Str$(Val(wert1$) Or Val(wert2$))
    Case 2
        erg$ = Str$(Val(wert1$) And Val(wert2$))
    Case 3
        erg$ = Str$(Val(wert1$) - Val(wert2$))
    Case 4
        erg$ = Str$(Val(wert1$) + Val(wert2$))
    Case 5
        If Val(wert2$) = 0 Then
            PkErrTxt$ = b$ + Chr$(0) + " -> Dividend Null!"
            PkErr% = True
            Call DefErrPop: Exit Function
        End If
        If Val(wert1$) = 0 Then
            erg$ = "0"
        Else
            erg$ = Str$(Val(wert1$) / Val(wert2$))
        End If
    Case 6
        erg$ = Str$(Val(wert1$) * Val(wert2$))
    Case 7
        erg$ = Str$((InStr(wert2$, wert1$) > 0))
    Case 10
        erg$ = Str$(Val(wert1$) < Val(wert2$))
    Case 20
        erg$ = Str$(Val(wert1$) = Val(wert2$))
    Case 30
        erg$ = Str$(Val(wert1$) > Val(wert2$))
    Case 12, 21
        erg$ = Str$(Val(wert1$) <= Val(wert2$))
    Case 13, 31
        erg$ = Str$(Val(wert1$) <> Val(wert2$))
    Case 23, 32
        erg$ = Str$(Val(wert1$) >= Val(wert2$))
    Case Else
    If minop% <> 99 Then Stop
    
    'Variable oder Konstante
    Select Case b$
        Case "WG"
            erg$ = PkWg$
        Case "AEP"
            erg$ = Str$(PkAep#)
        Case "KP"
            erg$ = Str$(PkKp#)
        Case "AVP"
            erg$ = Str$(PkAvp#)
        Case "MW"
            erg$ = Str$(PkMw!)
        Case "LC"
            erg$ = PkLc$
        Case "KAS"
            erg$ = PkKas$
        Case "KASZ"
            erg$ = PkKasz$
        Case "REZ"
            erg$ = PkRez$
        Case "AG"
            erg$ = Str$(PkAg)
        Case "AVPMIN"
            erg$ = Str$(PkAvpMin#)
        Case "AVPMAX"
            erg$ = Str$(PkAvpMax#)
        Case Else
            numerisch% = True
            For i% = 1 To Len(b$)
                If InStr("1234567890,.", Mid$(b$, i%, 1)) = 0 Then numerisch% = 0
            Next i%
            If (numerisch% = 0) Then
                If Left$(b$, 1) = Chr$(34) And Right$(b$, 1) = Chr$(34) Then
                    'StringKonstante
                    erg$ = Mid$(b$, 2, Len(b$) - 2)
                Else
                    PkErrTxt$ = b$ + Chr$(0) + " -> Variable nicht bekannt!"
                    PkErr% = True
                    Call DefErrPop: Exit Function
                End If
            Else
                'Komma auf Punkt tauschen
                Do
                  i% = InStr(b$, ","): If i% > 0 Then Mid$(b$, i%, 1) = "."
                Loop While (i% > 0)
                erg$ = b$
            End If
    End Select
End Select

PKparse$ = erg$

Call DefErrPop
End Function

Function PKfuparam$(kette$, ParamNr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PKfuparam$")
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
Dim p%, klammer%, ParStart%, AktPar%
Dim b$, b1$, c$

PKfuparam$ = ""

b$ = LTrim$(kette$)
b1$ = ""
'Klammern entfernen
If (Left$(b$, 1) = "(") And (Right$(b$, 1) = ")") Then
    p% = 1: klammer% = 0
    Do
        c$ = Mid$(b$, p%, 1)
        If (c$ = "(") Then
            klammer% = klammer% + 1
        ElseIf (c$ = ")") Then
            klammer% = klammer% - 1
            If (klammer% = 0) Then
                If (p% = Len(b$)) Then b$ = Mid$(b$, 2, Len(b$) - 2)
                Exit Do
            End If
        End If
        p% = p% + 1
    Loop While (p% <= Len(b$))
End If

'Parameter suchen
ParStart% = 1
AktPar% = 0
p% = 1: klammer% = 0
Do
    c$ = Mid$(b$, p%, 1)
    If (c$ = "(") Then
        klammer% = klammer% + 1
    ElseIf (c$ = ")") Then
        klammer% = klammer% - 1
    ElseIf (c$ = ",") Then
        If (klammer% = 0) Then
            AktPar% = AktPar% + 1
            If (AktPar% = ParamNr%) Then
                b1$ = Mid$(b$, ParStart%, p% - ParStart%)
                PKfuparam$ = PKparse$(b1$)
            Else
                ParStart% = p% + 1
            End If
        End If
    End If
    p% = p% + 1
Loop While (p% <= Len(b$))

If (b1$ = "") And (PkErr% = 0) Then
    If (ParamNr% = AktPar% + 1) Then
        b1$ = Mid$(b$, ParStart%, p% - ParStart%)
        PKfuparam$ = PKparse$(b1$)
    Else
        PkErr% = True
        PkErrTxt$ = b$ + Chr$(0) + "Parameter" + Str$(ParamNr%) + " nicht vorhanden!"
    End If
End If

Call DefErrPop
End Function


Sub PkBelegen(aep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PkBelegen")
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

With FlxPreisKorrektur
    pzn$ = .TextMatrix(.row, 0)
End With

FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ast.GetRecord (FabsRecno& + 1)
End If

PkAep# = aep#
PkKp# = ast.kp
PkAvp# = ast.AVP
PkMw! = para.Mwst(Val(ast.mw))
PkWg$ = ast.wg
PkLc$ = ast.Lac
PkKas$ = ast.kas
PkKasz$ = ast.kasz
PkRez$ = ast.rez
PkAg$ = Asc(ast.frei)

PkAvpMin# = 0
PkAvpMax# = 0
FabsErrf% = ZusWv3.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ZusWv3.GetRecord (FabsRecno& + 1)
    PkAvpMin# = ZusWv3.AvpMin
    PkAvpMax# = ZusWv3.AvpMax
End If

Call DefErrPop
End Sub

Function SelectPk%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SelectPk%")
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
Dim fPk%, z%, t%, ind%
Dim Zeile$(20), zStr$, h$

SelectPk% = -1

With frmEdit
    .Left = FlxPreisKorrektur.Left + FlxPreisKorrektur.ColPos(4) + 45
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = FlxPreisKorrektur.Top + (FlxPreisKorrektur.row - FlxPreisKorrektur.TopRow + 1) * FlxPreisKorrektur.RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
End With
With frmEdit.flxEdit
    .Left = 0
    .Top = 0
    
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Kalk-Name|>Kalk-Nr"
    .Rows = 1
    .Cols = 2
    
    Err = 0
    fPk% = FreeFile
    'On Error Resume Next
    Open "PKSCHEMA.TXT" For Input As #fPk%
    If Err = 0 Then
        While Not EOF(fPk%)
            Line Input #fPk%, zStr$
            Call SeparateKomma(zStr$)
            If (Val(argv$(2)) = 0) And (Val(argv$(3)) = 0) Then
                argv$(4) = Mid$(argv$(4), 2, Len(argv$(4)) - 2)
                If (Left$(argv$(4), 1) = "*") Then
                    h$ = Mid$(argv$(4), 2) + vbTab + argv$(1)
                    .AddItem h$
                End If
            End If
        Wend
        Close #fPk%
    End If
    
    .ColWidth(0) = frmEdit.TextWidth("Xxxxxxxxxxxxxxx Xxxxxxxxx Xxxxxxxxxxxxx")
    .ColWidth(1) = frmEdit.TextWidth("99999999")

    
    .Height = .RowHeight(0) * .Rows + 90
    .Width = .ColWidth(0) + .ColWidth(1) + 90
    
    frmEdit.Height = .Height
    frmEdit.Width = .Width
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

frmEdit.Show 1

If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    SelectPk% = Val(Mid$(EditTxt$, ind% + 1))
End If

Call DefErrPop
End Function

Sub AufschlagAuswahl()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AufschlagAuswahl")
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
Dim i%, Aufschlag%, auf%, mw%, ind%
Dim aep#, AVP#, AvpN#
Dim pzn$, h$

With FlxPreisKorrektur
    pzn$ = .TextMatrix(.row, 0)
    aep# = xVal(.TextMatrix(.row, 4))
    AvpN# = xVal(.TextMatrix(.row, 5))
End With

FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ast.GetRecord (FabsRecno& + 1)
    mw% = Val(ast.mw)
    Aufschlag% = Val(ast.ka)
Else
    mw% = 2
    Aufschlag% = 0
End If

'''''''''
With frmEdit
    .Left = FlxPreisKorrektur.Left + FlxPreisKorrektur.ColPos(4) + 45
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = FlxPreisKorrektur.Top + (FlxPreisKorrektur.row - FlxPreisKorrektur.TopRow + 1) * FlxPreisKorrektur.RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
End With
With frmEdit.flxEdit
    .Left = 0
    .Top = 0
    
    .Rows = 2
    .FixedRows = 1
    .FormatString = "^Aufschlag zu " + Format(aep#, "0.00") + "|>Avp|>Aufschlag-Nr"
    .Rows = 1
    .Cols = 3
    
    
    .ColWidth(0) = frmEdit.TextWidth("99999999999999.99")
    .ColWidth(1) = frmEdit.TextWidth("9999999999")
    .ColWidth(2) = 0

    For i% = 1 To 9
        auf% = para.Aufschlag(i%)
        AVP# = FNX(aep# * (1# + auf% / 100#) * (100# + para.Mwst(mw%)) / 100#)
        h$ = Format(auf%, "0") + "%" + vbTab
        h$ = h$ + Format(AVP#, "0.00") + vbTab + Format(i%, "0")
        .AddItem h$
    Next i%
    
    .Height = .RowHeight(0) * .Rows + 90
    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
    
    frmEdit.Height = .Height
    frmEdit.Width = .Width
    
    .row = 1
    If (Aufschlag% > 0) Then .row = Aufschlag%
    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

frmEdit.Show 1

If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    EditTxt$ = Mid$(EditTxt$, ind% + 1)
    ind% = InStr(EditTxt$, vbTab)
    AVP# = xVal(Left$(EditTxt$, ind% - 1))
    Aufschlag% = Val(Mid$(EditTxt$, ind% + 1))
    
    With FlxPreisKorrektur
        .TextMatrix(.row, 5) = Format(AVP#, "0.00")
    End With

    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        ast.ka = Format(Aufschlag%, "0")
        ast.PutRecord (FabsRecno& + 1)
    End If
    
    Call PreisKorrekturSpeichern
End If

Call DefErrPop
End Sub


Sub PreisKorrekturSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisKorrekturSpeichern")
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
Dim i%, Aufschlag%, auf%, mw%, ind%, rInd%, OldRow%, OrgRedraw%, row%
Dim aep#, AVP#, AvpN#
Dim pzn$, h$

With FlxPreisKorrektur
    aep# = xVal(.TextMatrix(.row, 4))
    AVP# = xVal(.TextMatrix(.row, 5))
    row% = xVal(.TextMatrix(.row, 10))
    
    With frmAction.flxarbeit(0)
        OrgRedraw% = .redraw
        .redraw = False
        OldRow% = .row
        .row = row%
    End With
    
    rInd% = SucheFlexZeile(False)
    If (rInd% > 0) Then
        ww.SatzLock (1)
        ww.WuAEP = aep#
        ww.WuAVP = AVP#
        ww.PutRecord (rInd% + 1)
        ww.SatzUnLock (1)
    End If
    
    With frmAction.flxarbeit(0)
        .redraw = OrgRedraw%
        .row = OldRow%
    End With
    
    PreisKalkAepChange% = True
End With

Call DefErrPop
End Sub


Sub WriteWZ(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WriteWZ")
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
Dim i%, fWz%, WzMax%, menge%, iHeute%, AblMonat%, AblJahr%, iAblauf%
Dim aep#
Dim KontrNr$, h$, heuteC$, pzn$, Meng2$, abl$
Dim WkZwi As clsWkZwi

h$ = Format(Now, "DDMMYY")
iHeute% = iDate(h$)

Set WkZwi = New clsWkZwi

WkZwi.OpenDatei

WkZwi.SatzLock (1)
WkZwi.GetFirstRecord
WzMax% = WkZwi.erstmax

pzn$ = ww.pzn
If (IstAltLast%) Then
    menge% = ww.LmAnzGebucht
Else
    menge% = ww.WuLm
End If

lif.GetRecord (ww.Lief + 1)

Meng2$ = Space$(5)

If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        Meng2$ = ast.herst
    End If
End If

With FlxPreisKorrektur
    KontrNr$ = Trim(.TextMatrix(pos%, 7))
    aep# = xVal(.TextMatrix(pos%, 4))
'    Call DxToMBFd(aep#)
    
    WkZwi.pzn = .TextMatrix(pos%, 0)
    WkZwi.datum = iHeute%
    WkZwi.Lief = Left$(Format(ww.Lief, "0") + Space$(3), 3) + lif.kurz
    WkZwi.menge = menge%
    WkZwi.KontrNr = Left$(KontrNr$ + Space$(11), 11)
    WkZwi.DruckKz = " "
    WkZwi.Meng2 = Meng2$
    WkZwi.Name = "  "
    WkZwi.IdDatum = iHeute%
    WkZwi.Preis = aep#
    WkZwi.Pruef = Space$(Len(WkZwi.Pruef))
    
    'Ablauf eintragen ------------------------------------------------------------
    iAblauf% = 0
    abl$ = RTrim(ww.WuAblDatum)
    If (Len(abl$) > 2) Then
        AblMonat% = Mid$(abl$, 3, 2)
        AblJahr% = Mid$(abl$, 5, 2)
        AblMonat% = AblMonat% + 1
        If (AblMonat% > 12) Then
            AblMonat% = 1
            AblJahr% = AblJahr% + 1
        End If
        abl$ = "01" + Format(AblMonat%, "00") + Format(AblJahr%, "00")
        iAblauf% = iDate(abl$) - 1
    End If
    WkZwi.AblaufDatum = iAblauf%
    
    WzMax% = WzMax% + 1
    WkZwi.PutRecord (WzMax% + 1)
End With

WkZwi.GetFirstRecord
WkZwi.erstmax = WzMax%
WkZwi.PutRecord (1)

WkZwi.SatzUnLock (1)
WkZwi.CloseDatei

Call DefErrPop
End Sub

'Sub WriteWZ(pos%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("WriteWZ")
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
'Dim i%, fWz%, WzMax%, menge%
'Dim aep#
'Dim KontrNr$, h$, heuteC$, pzn$, Meng2$
'
'h$ = Format(Now, "DDMMYY")
'heuteC$ = MKDate(iDate(h$))
'
'fWz% = FreeFile
'Open "EW-WKZWI.DAT" For Random Access Read Write Shared As #fWz% Len = 128
'
'Lock #fWz%, 1
'Get #fWz%, 1, WkZw
'WzMax% = CVI(Left$(WkZw.pzn, 2))
'
'pzn$ = ww.pzn
'If (IstAltLast%) Then
'    menge% = ww.LmAnzGebucht
'Else
'    menge% = ww.WuLm
'End If
'
'lif.GetRecord (ww.lief + 1)
'
'Meng2$ = Space$(5)
'
'If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
'    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        ast.GetRecord (FabsRecno& + 1)
'        Meng2$ = ast.herst
'    End If
'End If
'
'With FlxPreisKorrektur
'    KontrNr$ = Trim(.TextMatrix(pos%, 7))
'    aep# = xVal(.TextMatrix(pos%, 4))
'    Call DxToMBFd(aep#)
'
'    WkZw.pzn = .TextMatrix(pos%, 0)
'    WkZw.datum = heuteC$
'    WkZw.lief = Left$(Format(ww.lief, "0") + Space$(3), 3) + lif.kurz
'    WkZw.menge = menge%
'    WkZw.KontrNr = Left$(KontrNr$ + Space$(11), 11)
'    WkZw.DruKz = " "
'    WkZw.Meng2 = Meng2$
'    WkZw.Name = "  "
'    WkZw.IdDatum = heuteC$
'    WkZw.Preis = aep#
'    WkZw.Pruef = Space$(Len(WkZw.Pruef))
'
'    WzMax% = WzMax% + 1
'    Put #fWz%, WzMax% + 1, WkZw
'End With
'
'Get #fWz%, 1, WkZw
'WkZw.pzn = MKI(WzMax%) + Mid$(WkZw.pzn, 3)
'Put #fWz%, 1, WkZw
'
'Unlock #fWz%, 1
'Close #fWz%
'
'Call DefErrPop
'End Sub


