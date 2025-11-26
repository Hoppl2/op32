VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchHilfsTaxe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const INI_SECTION = "Auswahl-Hilfstaxe"
Const ARBEIT_SECTION = "Arbeitsbereich Auswahl-Hilfstaxe"
Const INFO_SECTION = "Infobereich Auswahl-Hilfstaxe"

Const EINGABE_NAME = "&Name/Pzn"

Dim MatchTaxeRec As Recordset

Dim FabsErrf%
Dim FabsRecno&

'''''''
Dim iMatchModus%
Dim iMatchTyp%
Dim iMatchListeTyp%
Dim iMatchAnzeigeTyp$()

Dim iMatchAutoRet%
Dim iMatchNurEiner%

Dim iOrgSuch$

Dim iMatchcodePzn$
Dim iMatchcodeTxt$
Dim iMatchcodeErg$
''''

Dim MatchKlein$, MatchGross$, MatchGueltig$


Private Const DefErrModul = "MATCHHILFSTAXE.CLS"

Sub InitMatch()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitMatch")
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
ReDim iMatchAnzeigeTyp$(0)

iMatchAnzeigeTyp$(0) = "Hilfstaxe"

MatchKlein = "abcdefghijklmnopqrstuvwxyzäöüß0123456789"
MatchGross = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜS0123456789"
MatchGueltig = MatchKlein + MatchGross
 
Call clsError.DefErrPop
End Sub

Function MakeMatchCode$(Name$, Optional mModus% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeMatchCode$")
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
Dim x$, ch$

x$ = Name$
'Call OemToChar(x$, x$)

i% = 1
While (i% <> 0)
    i% = InStr(x$, " ")
    If (i% <> 0) Then
        x$ = Left$(x$, i% - 1) + Mid$(x$, i% + 1)
    End If
Wend

For i% = 1 To Len(x$)
    ch$ = Mid$(x$, i%, 1)
    ch$ = Mid$(MatchGross$ + ch$, InStr(MatchKlein$ + ch$, ch$), -(InStr(MatchGueltig$, ch$) <> 0))
    Mid$(x$, i%, 1) = ch$
Next i%

'Call CharToOem(x$, x$)
MakeMatchCode$ = Left$(x$ + Space$(10), 10)

Call clsError.DefErrPop
End Function

Sub Umspeichern(buf$, pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Umspeichern")
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
Dim h$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

Ausgabe(pos%).pzn = Val(hTax1.pzn)
Ausgabe(pos%).Bookmark = hTax1.FabsRecno
'Ausgabe(pos%).key = clsfabs.LastKey(10)
Ausgabe(pos%).key = String("A", 10)


h$ = " " + vbTab + hTax1.kurz + vbTab
If (Para1.Land = "D") Then
    h$ = h$ + hTax1.meng
Else
    h$ = h$ + hTax1.Gstufe
End If
h$ = h$ + vbTab + hTax1.meh
If (Para1.Land = "D") Then
Else
    h$ = h$ + vbTab + Format(hTax1.kp, "0.00")
End If

'        dGstufe# = xVal(Ew.Gstufe)
'        If (dGstufe# = CInt(dGstufe#)) Then
'            meng$ = Format(dGstufe#, "0") + "    "
'        Else
'            meng$ = Format(dGstufe#, "0.000")
'        End If
        


'Call OemToChar(h$, h$)
Ausgabe(pos%).Name = h$
Ausgabe(pos%).Verweis = Len(h$)

Call clsError.DefErrPop
End Sub

Function SuchWeiter%(ind%, Typ%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchWeiter%")
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
Dim AltRecno&
Dim h$, AltKey$

SuchWeiter% = False

AltRecno& = Ausgabe(ind%).Bookmark
AltKey$ = Left$(RTrim$(Ausgabe(ind%).key) + Space$(10), 10)
If (Typ%) Then
    FabsErrf% = hTax1.IndexNext(1, AltRecno&, AltKey$, FabsRecno&)
Else
    FabsErrf% = hTax1.IndexPrevious(1, AltRecno&, AltKey$, FabsRecno&)
End If
If (FabsErrf% = 0) Then
    hTax1.GetRecord (FabsRecno& + 1)
    SuchWeiter% = True
End If

Call clsError.DefErrPop
End Function

Function SuchArtikel%(Such$, iArbeitAnzzeilen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchArtikel%")
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
Dim i%, l%, l2%, ind%, TextSuche%, IstPzn%, LeftLen%, InhMengeFlag%, gef%, IstCode%
Dim InhMenge&
Dim SQLStr$, nSuch$, ch$
Dim iBookmark As Variant

SuchArtikel% = True

iOrgSuch$ = Such$

If (Such$ = "") Then Such$ = Space$(10)
l% = Len(Such$)
If (l% = 8) And (Left$(Such$, 1) = "-") Then Such$ = Mid$(Such$, 2)
If ((l% >= 10) And (((Left$(Such$, 1) = "*") And (Mid$(Such$, l%, 1) = "*")) Or (InStr(Such$, "PZN") > 0))) Then
    nSuch$ = ""
    For i% = 1 To l%
        ch$ = Mid$(Such$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            nSuch$ = nSuch$ + ch$
        End If
    Next i%
    Such$ = nSuch$
End If

l% = Len(Such$)

IstCode% = False
If (l% <= 7) Then
    IstCode% = True
    For i% = 1 To l%
        ch$ = Mid$(Such$, i%, 1)
        If (InStr("0123456789", ch$) <= 0) Then
            IstCode% = False
            Exit For
        End If
    Next i%
End If

If (IstCode%) Then
    nSuch$ = Right$(Space$(7) + Such$, 7)
    
    FabsErrf% = hTax1.IndexSearch(0, Such$, FabsRecno&)
    If (FabsErrf% = 0) Then
        If (iMatchNurEiner% = False) Then
            hTax1.GetRecord (FabsRecno& + 1)
            Call SuchArtikelTrue(True)
        End If
    Else
        SuchArtikel% = False
    End If
    
    If (SuchArtikel% And iMatchNurEiner%) Then
        iMatchcodePzn$ = Format$(Such$, "0000000")
        iMatchcodeErg$ = iMatchcodePzn$
        Unload frmMatchcode
    End If

    Call clsError.DefErrPop: Exit Function
End If

Such$ = MakeMatchCode$(Such$)
FabsErrf% = hTax1.IndexGeneric(1, Such$, FabsRecno&)
If ((FabsErrf% = 0) Or (FabsErrf% = 12) Or (FabsErrf% = 13) Or (FabsErrf% = 15)) Then
    hTax1.GetRecord (FabsRecno& + 1)
    Call SuchArtikelTrue(False)
Else
    SuchArtikel% = False
End If

Call clsError.DefErrPop
End Function

'Sub SuchArtikelTrue(Optional PznFlag% = False)
Sub SuchArtikelTrue(PznFlag%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchArtikelTrue")
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
Dim j%, ind%
Dim h$

Call Umspeichern("", 0)
Call frmMatchcode.MachAuswahlGrid
If (PznFlag%) Then
    h$ = Ausgabe(0).Name
    With frmMatchcode.flxarbeit(0)
        .Rows = 2
        For j% = 0 To 2
            ind% = InStr(h$, vbTab)
            .TextMatrix(1, j%) = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Next j%
        .TextMatrix(1, 3) = h$
        .row = 1
    End With
    Call ListeAusFlexBefuellen
    With frmMatchcode.txtMatchcode
        .SelStart = 0
        .SelLength = Len(.text)
    End With
Else
    With frmMatchcode
        Call .MachMaske
        .flxarbeit(0).Visible = True
        .txtFlexBack.Visible = True
        .flxarbeit(1).Visible = False
    End With
    With frmMatchcode.flxarbeit(0)
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
        .SetFocus
    End With
End If

Call clsError.DefErrPop
End Sub

Sub InitAuswahlGrid()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitAuswahlGrid")
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
Dim spBreite%, i%
Dim FormStr$

With frmMatchcode.flxarbeit(0)
            
    FormStr$ = "|<Name|>Menge|<Eh"
    If (Para1.Land = "A") Then
        FormStr$ = FormStr$ + "|>TaxAns"
    End If
    .FormatString = FormStr$
    
    frmMatchcode.Font.Bold = True
    .ColWidth(0) = frmMatchcode.TextWidth("XX")
    .ColWidth(2) = frmMatchcode.TextWidth("XXXXXXX")
    .ColWidth(3) = frmMatchcode.TextWidth("XXXX")
    If (Para1.Land = "A") Then
        .ColWidth(4) = frmMatchcode.TextWidth(String(8, "9"))
    End If
    frmMatchcode.Font.Bold = False
    
End With

Call clsError.DefErrPop
End Sub

Sub ListeAusFlexBefuellen(Optional AlphaText$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeAusFlexBefuellen")
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
Dim row%
Dim pzn$
        
If (AlphaText$ <> "") Then
    Call ListeBefuellen("9999999", AlphaText$, "", "")
Else
    
    With frmMatchcode.flxarbeit(0)
        row% = .row
        
        pzn$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
        Call ListeBefuellen(pzn$, .TextMatrix(row%, 1), .TextMatrix(row%, 2), .TextMatrix(row%, 3))
    End With
End If

Call clsError.DefErrPop
End Sub

Sub ListeAusStiftBefuellen(Dp04Str$, Dp04Menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeAusStiftBefuellen")
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
Dim row%
Dim h$, pzn$, SQLStr$, AlphaText$
        
AlphaText$ = ""

h$ = Dp04Str$
h$ = Right$(Space$(13) + h$, 13)
FabsErrf% = Ass1.IndexSearch(1, h$, FabsRecno&)
If (FabsErrf% = 0) Then
    Ass1.GetRecord (FabsRecno& + 1)
    pzn$ = Ass1.pzn
Else
    If (Len(Dp04Str$) = 7) Then
        h$ = Dp04Str$
    Else
        h$ = Mid$(h$, 6, 7)
    End If
    FabsErrf% = Ass1.IndexSearch(0, h$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Ass1.GetRecord (FabsRecno& + 1)
        pzn$ = Ass1.pzn
    Else
        pzn$ = "9999999"
        AlphaText$ = "Artikel unbekannt"
    End If
End If

Call ListeBefuellen(pzn$, AlphaText$, "", "", Dp04Menge%)

Call clsError.DefErrPop
End Sub
  
Sub ListeBefuellen(pzn$, txt1$, txt2$, txt3$, Optional Dp04Menge% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeBefuellen")
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
Dim i%, erg%, row%, ManuellAsatz%, ManuellSsatz%, ManuellBm%, ManuellNm%, abl%, ind%, iMenge%
Dim h$, h2$, AufLager$
        
iMenge% = 1
If (Dp04Menge% > 0) Then iMenge% = Dp04Menge%

If (pzn$ = "9999999") Then
    h$ = "9999999" + vbTab + txt1$ + vbTab + vbTab + vbTab
    h$ = h$ + Str$(iMenge%)
Else
    ManuellAsatz% = 0
    FabsErrf% = hTax1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ManuellAsatz% = CInt(FabsRecno&)
        hTax1.GetRecord (FabsRecno& + 1)
        iMenge% = hTax1.meng
    End If
            
    If (Trim(txt1$) = "") Then
        txt1$ = hTax1.kurz
        txt2$ = hTax1.meng
        txt3$ = hTax1.meh
    End If
    
    h$ = pzn$ + vbTab + Trim$(txt1$) + vbTab + Trim$(txt2$) + vbTab + Trim$(txt3$)
    h$ = h$ + vbTab + Str$(iMenge%)
End If

With frmMatchcode.flxarbeit(1)
    .AddItem h$, 1
    .row = 1
    .col = 3
    .col = 4
    .ColSel = 4
    
    .Visible = True
    frmMatchcode.flxarbeit(0).Visible = False
    frmMatchcode.txtFlexBack.Visible = False
    
    frmMatchcode.txtMatchcode.SetFocus
 
End With

Call clsError.DefErrPop
End Sub

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, arow%, m%, aMeng%, iLief%
Dim h$, h2$
            
With frmMatchcode.flxarbeit(1)
    EditCol% = .col
    If ((EditCol% = 1) And (.TextMatrix(.row, 0) <> "9999999")) Or _
       ((iMatchListeTyp% = 0) And (EditCol% <> 4) And (EditCol% <> 5)) Or _
       ((iMatchListeTyp% = 1) And ((EditCol% = 2) Or (EditCol% = 3))) Then
        Call clsError.DefErrPop: Exit Sub
    End If
            
        
'    If ((EditCol% = 1) And (.TextMatrix(.row, 0) = "9999999")) Or (EditCol% = 4) Or (EditCol% = 5) Then
        
        KeinRowColChange% = True
        arow% = .row
        .row = 0
        .CellFontBold = True
        .row = arow%
        KeinRowColChange% = False
            
        Load frmEdit2
        
        If (iMatchListeTyp% = 1) And (EditCol% = 4) Then
            lInd% = -1
            With frmMatchcode
                frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(EditCol%) + 45
                frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + .flxarbeit(1).RowHeight(0)
                frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                frmEdit2.Width = .flxarbeit(1).ColWidth(EditCol%) * 1.5
                frmEdit2.Height = .flxarbeit(1).Height - .flxarbeit(1).RowHeight(0)
            End With
            With frmEdit2.lstEdit
                .Height = frmEdit2.ScaleHeight
                frmEdit2.Height = .Height
                .Width = frmEdit2.ScaleWidth
                .Left = 0
                .Top = 0
                
                .Clear
                If (LieferantenDBok) Then
                    For i% = 1 To LieferantenDB1.AnzLiefNamen
                        h$ = LieferantenDB1.LiefName(i% - 1)
                        .AddItem h$
                    Next i%
                Else
                    For i% = 1 To Lif1.AnzLiefNamen
                        h$ = Lif1.LiefName(i% - 1)
                        .AddItem h$
                    Next i%
                End If
                .Visible = True
                EditModus% = 1
            End With
            
        Else
        
            With frmMatchcode
                frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(EditCol%) + 45
                frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + (.flxarbeit(1).row - .flxarbeit(1).TopRow + 1) * .flxarbeit(1).RowHeight(0)
                frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                frmEdit2.Width = .flxarbeit(1).ColWidth(EditCol%)
                frmEdit2.Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
            End With
            With frmEdit2.txtEdit
                .Width = frmEdit2.ScaleWidth
    '            .Height = frmEdit.ScaleHeight
                .Left = 0
                .Top = 0
                h2$ = frmMatchcode.flxarbeit(1).TextMatrix(frmMatchcode.flxarbeit(1).row, EditCol%)
                .text = h2$
                .BackColor = vbWhite
                .Visible = True
            End With
            
            EditModus% = 0
            If (EditCol% = 1) Then
                EditModus% = 1
            ElseIf (iMatchListeTyp% = 1) Then
                If (EditCol% = 5) Then
                    EditModus% = 1
                ElseIf (EditCol% = 6) Then
                    EditModus% = 6
                ElseIf (EditCol% = 8) Then
                    EditModus% = 5
                ElseIf (EditCol% = 9) Then
                    EditModus% = 4
                End If
            End If
        End If
            
        frmEdit2.Show 1
            
        KeinRowColChange% = True
        arow% = .row
        .row = 0
        .CellFontBold = False
        .row = arow%
        KeinRowColChange% = False
            
        If (EditErg%) Then
            If (EditModus% = 0) Then
                m% = Val(EditTxt$)
                .TextMatrix(.row, EditCol%) = Mid$(Str$(m%), 2)
            ElseIf (EditModus% = 1) Or (EditModus% = 5) Or (EditModus% = 6) Then
                .TextMatrix(.row, EditCol%) = Trim(EditTxt$)
            ElseIf (EditModus% = 4) Then
                .TextMatrix(.row, EditCol%) = Format(Val(EditTxt$), "0.00")
            End If
        End If
'    End If
End With

Call clsError.DefErrPop
End Sub
      

Public Property Get IniSection$()
IniSection$ = INI_SECTION
End Property

Public Property Get InfoSection$()
InfoSection$ = INFO_SECTION
End Property

Public Property Get ArbeitSection$()
ArbeitSection$ = ARBEIT_SECTION
End Property

Public Property Get OrgCols%()
If (Para1.Land = "D") Then
    OrgCols% = 4
Else
    OrgCols% = 5
End If
End Property

Public Property Get EingabeName$()
EingabeName$ = EINGABE_NAME
End Property



Public Property Get MatchModus%()
MatchModus% = iMatchModus%
End Property
Public Property Let MatchModus(ByVal vNewValue%)
iMatchModus% = vNewValue%
End Property


Public Property Get MatchTyp%()
MatchTyp% = iMatchTyp%
End Property
Public Property Let MatchTyp(ByVal vNewValue%)
iMatchTyp% = vNewValue%
End Property


Public Property Get MatchListeTyp%()
MatchListeTyp% = iMatchListeTyp%
End Property
Public Property Let MatchListeTyp(ByVal vNewValue%)
iMatchListeTyp% = vNewValue%
End Property


Public Property Get MatchAnzeigeTyp$(ByVal ind%)
MatchAnzeigeTyp$ = iMatchAnzeigeTyp$(ind%)
End Property
Public Property Let MatchAnzeigeTyp(ByVal ind%, ByVal vNewValue$)
iMatchAnzeigeTyp$(ind%) = vNewValue$
End Property


Public Property Get MatchAutoRet%()
MatchAutoRet% = iMatchAutoRet%
End Property
Public Property Let MatchAutoRet(ByVal vNewValue%)
iMatchAutoRet% = vNewValue%
End Property


Public Property Get MatchNurEiner%()
MatchNurEiner% = iMatchNurEiner%
End Property
Public Property Let MatchNurEiner(ByVal vNewValue%)
iMatchNurEiner% = vNewValue%
End Property


Public Property Get OrgSuch$()
OrgSuch$ = iOrgSuch$
End Property
Public Property Let OrgSuch(ByVal vNewValue$)
iOrgSuch$ = vNewValue$
End Property


Public Property Get MatchcodePzn$()
MatchcodePzn$ = iMatchcodePzn$
End Property
Public Property Let MatchcodePzn(ByVal vNewValue$)
iMatchcodePzn$ = vNewValue$
End Property


Public Property Get MatchcodeTxt$()
MatchcodeTxt$ = iMatchcodeTxt$
End Property
Public Property Let MatchcodeTxt(ByVal vNewValue$)
iMatchcodeTxt$ = vNewValue$
End Property


Public Property Get MatchcodeErg$()
MatchcodeErg$ = iMatchcodeErg$
End Property
Public Property Let MatchcodeErg(ByVal vNewValue$)
iMatchcodeErg$ = vNewValue$
End Property


Public Property Get UBoundMatchAnzeigeTyp%()
UBoundMatchAnzeigeTyp% = UBound(iMatchAnzeigeTyp$)
End Property
Public Property Let UBoundMatchAnzeigeTyp(ByVal vNewValue%)
'UBound(iMatchAnzeigeTyp$)= vNewValue%
End Property


Public Property Get InfoTyp%()
InfoTyp% = 0
End Property

Public Property Get WarteMaske%()
End Property
Public Property Let WarteMaske(ByVal vNewValue%)
End Property



