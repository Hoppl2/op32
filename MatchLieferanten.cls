VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchLieferanten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const INI_SECTION = "Auswahl-Lieferanten"
Const ARBEIT_SECTION = "Arbeitsbereich Auswahl-Lieferanten"
Const INFO_SECTION = "Infobereich Auswahl-Lieferanten"

Const EINGABE_NAME = "&Name/Nr"

Dim MatchTaxeRec As New ADODB.Recordset

Dim LiefNr%
Dim FabsRecno&

Dim oMatchModus%

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



Private Const DefErrModul = "MATCHLIEFERANTEN.CLS"

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
ReDim iMatchAnzeigeTyp$(1)
iMatchAnzeigeTyp$(0) = "Lieferanten"
iMatchAnzeigeTyp$(1) = "Lieferanten numerisch"

oMatchModus% = -1

Call clsError.DefErrPop
End Sub

Sub InitLst()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitLst")
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
Dim i%, j%
Dim h$

With frmLiefMatchcode.lstSortierung
    .Clear
    j% = 0
    For i% = 1 To Lif1.AnzRec   ' 200
        Call Lif1.GetRecord(i% + 1)
        h$ = Lif1.kurz
        h$ = UCase(Trim$(h$))
        If (h$ <> "") Then
            If (Asc(Left$(h$, 1)) >= 32) Then
                If (iMatchModus% = LIEFERANTEN_ALPHA) Then
                    h$ = h$ + " (" + Mid$(Str$(i%), 2) + ")"
                Else
                    h$ = Format(i%, "000") + " (" + Mid$(Str$(i%), 2) + ")"
                End If
                .AddItem h$
            End If
        End If
    Next i%
End With
    
Call clsError.DefErrPop
End Sub

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
Dim ind%, ind2%, i%, z%, z1%, z2%
Dim zu!
Dim EK#, vk#
Dim h$, h2$, char$, zuza$, zusa$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

If (LieferantenDBok) Then
    Ausgabe(pos%).pzn = MatchTaxeRec!LiefNr
    
    h$ = " "
    Ausgabe(pos%).Name = h$ + vbTab + MachTaxeAuswahlName$(MatchTaxeRec)
    Ausgabe(pos%).Verweis = Len(RTrim(Ausgabe(pos%).Name))
    
    If (MatchTaxeRec!LiefNr < 10) Then Ausgabe(pos%).LagerKz = 2
    Ausgabe(pos%).Bookmark = MatchTaxeRec.Bookmark
Else
    Ausgabe(pos%).pzn = LiefNr%
    Ausgabe(pos%).Bookmark = FabsRecno&
    
    h$ = " "
    FabsErrf% = lZus1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "00000") + "001", FabsRecno&)
    If (FabsErrf% = 0) Then
        Ausgabe(pos%).ZusatzTextKz = 2
        h$ = "+"
    End If
    
    If (LiefNr% < 10) Then Ausgabe(pos%).LagerKz = 2
    
'        h$ = h$ + vbTab + Lif1.kurz + vbTab + Str$(LiefNr%) + vbTab + Lif1.Name(0) + vbTab + Lif1.Name(1) + vbTab
'        h$ = h$ + Lif1.telefon + vbTab + Lif1.Durchwahl
    
    h2$ = RTrim$(Lif1.kurz)
    h$ = h$ + vbTab + h2$ + vbTab + Mid$(Str$(LiefNr%), 2)
    
    Ausgabe(pos%).Name = h$
    Ausgabe(pos%).Verweis = Len(h$)
End If
        
Call clsError.DefErrPop
End Sub

Function MachTaxeAuswahlName$(tRec As ADODB.Recordset)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachTaxeAuswahlName$")
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
Dim ret$

ret$ = clsOpTool.CheckNullStr(tRec!kurz) + vbTab + Format(tRec!LiefNr, "0")
MachTaxeAuswahlName$ = ret$

Call clsError.DefErrPop
End Function

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
Dim erg%, ind2%
Dim AltRecno&
Dim h$

SuchWeiter% = False

erg% = False

If (LieferantenDBok) Then
    MatchTaxeRec.Bookmark = Ausgabe(ind%).Bookmark
    Do
        If (Typ%) Then
            MatchTaxeRec.MoveNext
            If (MatchTaxeRec.EOF) Then
                Exit Do
            Else
'                If (AlleKunden%) Or (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
                    SuchWeiter% = True
                    Exit Do
'                End If
            End If
        Else
            MatchTaxeRec.MovePrevious
            If (MatchTaxeRec.BOF) Then
                Exit Do
            Else
'                If (AlleKunden%) Or (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
                    SuchWeiter% = True
                    Exit Do
'                End If
            End If
        End If
    Loop
Else
    With frmLiefMatchcode.lstSortierung
        AltRecno& = Ausgabe(ind%).Bookmark
        If (Typ%) Then
            If (AltRecno& < (.ListCount - 1)) Then
                FabsRecno& = AltRecno& + 1
                erg% = True
            End If
        Else
            If (AltRecno& > 0) Then
                FabsRecno& = AltRecno& - 1
                erg% = True
            End If
        End If
        
        If (erg%) Then
            .ListIndex = FabsRecno&
            h$ = .text
            ind2% = InStr(h$, "(")
            h$ = Mid$(h$, ind2% + 1)
            
            ind2% = Val(Left$(h$, Len(h$) - 1))
            LiefNr% = ind2%
        
            Lif1.GetRecord (LiefNr% + 1)
            SuchWeiter% = True
        End If
    End With
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
Dim i%, ind%
Dim h$, SQLStr$

SuchArtikel% = True

iOrgSuch$ = Such$

'''''''''''
If (LieferantenDBok) Then
    iMatchModus% = LIEFERANTEN_NUMERISCH
    If (Such$ = "") Or (Len(Such$) > 3) Then
        iMatchModus% = LIEFERANTEN_ALPHA
    Else
        For i% = 1 To Len(Such$)
            If (InStr("0123456789", Mid$(Such$, i%, 1)) = 0) Then
                iMatchModus% = LIEFERANTEN_ALPHA
                Exit For
            End If
        Next i%
    End If
    
'    Set MatchTaxeRec = LieferantenDB.OpenRecordset("Lieferanten", dbOpenTable)
'    If (iMatchModus% = LIEFERANTEN_NUMERISCH) Then
'        MatchTaxeRec.Index = "LiefNr"
'    Else
'        MatchTaxeRec.Index = "Kurz"
'    End If
'    MatchTaxeRec.Seek ">=", Such$
'    If (MatchTaxeRec.NoMatch = True) Then
'        SuchArtikel% = False
'    Else
''        Call Umspeichern("", 0)
'    End If
    
    On Error Resume Next
    MatchTaxeRec.Close
    Err.Clear
    On Error GoTo DefErr

    SQLStr = "SELECT * FROM Lieferanten WHERE "
    If (iMatchModus% = LIEFERANTEN_NUMERISCH) Then
        SQLStr = SQLStr + "LiefNr>=" + Such$
        SQLStr = SQLStr + " ORDER BY LiefNr"
    Else
        SQLStr = SQLStr + "Kurz>='" + Such$ + "'"
        SQLStr = SQLStr + " ORDER BY Kurz"
    End If
    MatchTaxeRec.Open SQLStr, LieferantenDB1.ActiveConn ' LieferantenConn
    If (MatchTaxeRec.EOF) Then
        SuchArtikel% = False
    Else
'        Call Umspeichern("", 0)
    End If
    
    If (SuchArtikel% = True) Then
        Call SuchArtikelTrue(0)
    End If
Else
    iMatchModus% = LIEFERANTEN_NUMERISCH
    If (Such$ = "") Or (Len(Such$) > 3) Then
        iMatchModus% = LIEFERANTEN_ALPHA
    Else
        For i% = 1 To Len(Such$)
            If (InStr("0123456789", Mid$(Such$, i%, 1)) = 0) Then
                iMatchModus% = LIEFERANTEN_ALPHA
                Exit For
            End If
        Next i%
    End If
    
    
    If (oMatchModus% <> iMatchModus%) Then
        Call InitLst
        oMatchModus% = iMatchModus%
    End If
    
    If (iMatchModus% = LIEFERANTEN_NUMERISCH) Then
        Such$ = Format(Val(Such$), "000")
    End If
    
    With frmLiefMatchcode.lstSortierung
        FabsRecno& = 0
        For i% = 1 To .ListCount
            .ListIndex = i% - 1
            h$ = .text
            If (h$ >= Such$) Then
                FabsRecno& = i% - 1
                Exit For
            End If
        Next i%
        
        .ListIndex = FabsRecno&
        h$ = .text
        ind% = InStr(h$, "(")
        h$ = Mid$(h$, ind% + 1)
        
        ind% = Val(Left$(h$, Len(h$) - 1))
        LiefNr% = ind%
    End With
    
    
    Lif1.GetRecord (LiefNr% + 1)
    Call Umspeichern("", 0)
    Call frmLiefMatchcode.MachAuswahlGrid
    'Call MachAuswahlGrid(iArbeitAnzzeilen%)
    Call frmLiefMatchcode.MachMaske
    With frmLiefMatchcode.flxarbeit(0)
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
        .SetFocus
    End With
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

Sub MachAuswahlGrid(iArbeitAnzzeilen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachAuswahlGrid")
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
Dim aFont$

With frmMatchcode.flxarbeit(0)
    .Rows = iArbeitAnzzeilen%
    .FixedRows = 1
            
    .Cols = 7
    .FormatString = "|<Kurz|>Nr|<Name1|<Name2|<Telefon|<Dw"
    frmLiefMatchcode.Font.Bold = True
    .ColWidth(0) = frmLiefMatchcode.TextWidth("X")
    .ColWidth(1) = frmLiefMatchcode.TextWidth("XXXXXXX")
    .ColWidth(2) = frmLiefMatchcode.TextWidth("999")
    .ColWidth(4) = frmLiefMatchcode.TextWidth(String$(25, "X"))
    .ColWidth(5) = frmLiefMatchcode.TextWidth(String$(16, "9"))
    .ColWidth(6) = frmLiefMatchcode.TextWidth(String$(5, "9"))
    frmLiefMatchcode.Font.Bold = False
            
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (i <> 3) Then
            .ColWidth(i%) = .ColWidth(i%) + frmLiefMatchcode.TextWidth("X")
            spBreite% = spBreite% + .ColWidth(i%)
        End If
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(3) = .Width - spBreite%
End With
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
Dim aFont$

With frmMatchcode.flxarbeit(0)
    .Cols = 3
    .FormatString = "|<Kurz|>Nr"
    frmMatchcode.Font.Bold = True
    .ColWidth(0) = frmMatchcode.TextWidth("XX")
    .ColWidth(1) = frmMatchcode.TextWidth("XXXXXXX")
    .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
    frmMatchcode.Font.Bold = False
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
OrgCols% = 3
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
InfoTyp% = 4
End Property


Public Property Get WarteMaske%()
End Property
Public Property Let WarteMaske(ByVal vNewValue%)
End Property



