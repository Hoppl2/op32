VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatchKundenADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const INI_SECTION = "Auswahl-Kunden"
Const ARBEIT_SECTION = "Arbeitsbereich Auswahl-Kunden"
Const INFO_SECTION = "Infobereich Auswahl-Kunden"

Const EINGABE_NAME = "&Name/Nr"

Dim MatchTaxeRec As New ADODB.Recordset

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

Dim KundenMaske$

'Dim OrgSuchBegriffe$(10), OrgSb$(10)
'Dim AnzSuchBegriffe%


Private Const DefErrModul = "MATCHKUNDENADO.CLS"

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

If (iMatchTyp = MATCH_KUNDEN_MEDVW) Then
    ReDim iMatchAnzeigeTyp$(1)
    iMatchAnzeigeTyp$(0) = "Kunden Medikations-Verwaltung"
    iMatchAnzeigeTyp$(1) = "Kunden"
Else
    ReDim iMatchAnzeigeTyp$(0)
    iMatchAnzeigeTyp$(0) = "Kunden"
End If

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
Dim h$, h2$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

If (KundenMdbOk%) Then
    Ausgabe(pos%).pzn = MatchTaxeRec!KundenNr
    
    h$ = " "
    Ausgabe(pos%).Name = h$ + vbTab + MachTaxeAuswahlName$(MatchTaxeRec)
    Ausgabe(pos%).Verweis = Len(RTrim(Ausgabe(pos%).Name))
    
    Ausgabe(pos%).LagerKz = Abs(clsOpTool.CheckNullLong(MatchTaxeRec!Status))
    Ausgabe(pos%).Bookmark = MatchTaxeRec.Bookmark
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

ret$ = clsOpTool.CheckNullStr(tRec!kurz) + vbTab + Format(tRec!KundenNr, "0")
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
Dim AlleKunden%
Dim AltRecno&
Dim h$, AltKey$

SuchWeiter% = False

AlleKunden% = (frmMatchcode.chkMatchcode.Value = 0)

If (KundenMdbOk%) Then
    MatchTaxeRec.Bookmark = Ausgabe(ind%).Bookmark
    Do
        If (Typ%) Then
            MatchTaxeRec.MoveNext
            If (MatchTaxeRec.EOF) Then
                Exit Do
            Else
                If (AlleKunden%) Or (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
                    SuchWeiter% = True
                    Exit Do
                End If
            End If
        Else
            MatchTaxeRec.MovePrevious
            If (MatchTaxeRec.BOF) Then
                Exit Do
            Else
                If (AlleKunden%) Or (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
                    SuchWeiter% = True
                    Exit Do
                End If
            End If
        End If
    Loop
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
Dim i%, j%, l%, ind%, gef%, IstCode%, IstPzn%, TextSuche%
Dim IstKundenNr%, IstKurz%, IstMaske%, AnzSuchTeile%, AlleKunden%
Dim KundenNr&, KundenMax&
Dim h$, SQLStr$, nSuch$, ch$, AltKey$, SuchTeil$, SuchName$(20), SuchOp$(20), SuchWert$(20)
Dim iBookmark As Variant

SuchArtikel% = True

iOrgSuch$ = Such$

If (KundenMdbOk%) Then
    IstKundenNr% = 0
    IstKurz% = 0
    IstMaske% = 0
    
    Do
        ind = InStr(Such, "*")
        If (ind > 0) Then
            Such = Left(Such, ind - 1) + "%" + Mid(Such, ind + 1)
        Else
            Exit Do
        End If
    Loop
    
    nSuch$ = Such$
    l = Len(nSuch)
    If (Right$(nSuch, 1) <> vbTab) Then
        IstKundenNr% = True
        For i% = 1 To l%
            ch$ = Mid$(nSuch$, i%, 1)
            If (InStr("0123456789", ch$) <= 0) Then
                IstKundenNr% = False
                Exit For
            End If
        Next i%
        
        If (IstKundenNr%) Then
            nSuch$ = "KUNDENNR,=," + nSuch + vbTab
        Else
            nSuch$ = "KURZ,=," + nSuch + vbTab
        End If
    End If
    For AnzSuchTeile% = 0 To 20
        ind% = InStr(nSuch$, vbTab)
        If (ind% > 0) Then
            SuchTeil$ = Left$(nSuch$, ind% - 1)
            nSuch$ = Mid$(nSuch$, ind% + 1)
            
            If (Right$(SuchTeil$, 1) <> ",") Then
                SuchTeil$ = SuchTeil$ + ","
            End If
            For i% = 0 To 2
                ind% = InStr(SuchTeil$, ",")
                If (ind% > 0) Then
                    h$ = Trim(Left$(SuchTeil$, ind% - 1))
                    If (i% = 0) Then
                        SuchName$(AnzSuchTeile%) = h$
                        If (UCase(h$) = "KUNDENNR") Then
                            IstKundenNr% = True
                        ElseIf (UCase(h$) = "KURZ") Then
                            IstKurz% = True
                        End If
                    ElseIf (i% = 1) Then
                        SuchOp$(AnzSuchTeile%) = h$
                    Else
                        SuchWert$(AnzSuchTeile%) = h$
                        For j% = 1 To Len(h$)
                            ch$ = Mid$(Such$, i%, 1)
                            If (InStr(",:*?<>", ch$) > 0) Then
                                IstMaske% = True
                                Exit For
                            End If
                        Next j%
                    End If
                    SuchTeil$ = Mid$(SuchTeil$, ind% + 1)
                Else
                    Exit For
                End If
            Next i%
        Else
            Exit For
        End If
    Next AnzSuchTeile%
    
    AlleKunden% = (frmMatchcode.chkMatchcode.Value = 0)

    If (iMatchTyp = MATCH_KUNDEN_MEDVW) Then
        If (IstKundenNr%) Then
            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr = " + SuchWert$(0)
            If (AlleKunden% = 0) Then
                SQLStr = SQLStr + " AND Status=0"
            End If
        Else
            SQLStr$ = "SELECT * FROM Kunden"
            For i% = 0 To (AnzSuchTeile% - 1)
                If (i% = 0) Then
                    SQLStr$ = SQLStr$ + " WHERE "
                Else
                    SQLStr$ = SQLStr$ + " AND "
                End If
                SQLStr$ = SQLStr$ + SuchName$(i%) + " " + SuchOp$(i%) + " '" + SuchWert$(i%) + "'"
            Next i%
            If (AlleKunden% = 0) Then
                SQLStr = SQLStr + " AND Status=0"
            End If
            If (iMatchModus = 0) Then
                SQLStr = SQLStr + " AND (MedVw<>0)"
            End If
            SQLStr$ = SQLStr$ + " ORDER BY Name, Vorname"
        End If
        
        On Error Resume Next
        MatchTaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        MatchTaxeRec.Open SQLStr$, KundenDB1.ActiveConn

        If (MatchTaxeRec.EOF) Then
            SuchArtikel% = False
        Else
        End If
        
        
'        SQLStr$ = "SELECT * FROM Kunden"
'        SQLStr$ = SQLStr$ + " WHERE Kurz>='" + SuchWert(0) + "'"
'        If (iMatchModus = 0) Then
'            SQLStr = SQLStr + " AND (MedVw<>0) AND (Status=0)"
'        End If
'        SQLStr$ = SQLStr$ + " ORDER BY Kurz, Name, Vorname"
'
'        Set MatchTaxeRec = KundenDB.OpenRecordset(SQLStr$)
'
'        If (MatchTaxeRec.EOF) Then
'            SuchArtikel% = False
'        Else
'        End If
    ElseIf (IstKurz%) And (IstMaske% = 0) Then
        SQLStr$ = "SELECT * FROM Kunden WHERE Kurz>='" + SuchWert$(0) + "'"
        If (AlleKunden = 0) Then
            SQLStr = SQLStr + " AND Status=0"
        End If
        SQLStr = SQLStr + " ORDER BY Kurz,Name,VorName,KundenNr"
        On Error Resume Next
        MatchTaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        MatchTaxeRec.Open SQLStr$, KundenDB1.ActiveConn
        If (MatchTaxeRec.EOF) Then
            SuchArtikel% = False
        End If
        
'        Set MatchTaxeRec = KundenDB.OpenRecordset("Kunden", dbOpenTable)
'        MatchTaxeRec.index = "Kurz"
'        MatchTaxeRec.Seek ">=", SuchWert$(0)
'        If (MatchTaxeRec.NoMatch = True) Then
'            SuchArtikel% = False
'        Else
'    '        Call Umspeichern("", 0)
'            If (AlleKunden% = 0) Then
'                SuchArtikel% = False
'                Do
'                    If (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
'                        SuchArtikel = True
'                        Exit Do
'                    End If
'                    MatchTaxeRec.MoveNext
'                    If (MatchTaxeRec.EOF) Then
'                        Exit Do
'                    End If
'                Loop
'            End If
'        End If
    Else
        If (IstKundenNr%) Then
            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr = " + SuchWert$(0)
            If (AlleKunden% = 0) Then
                SQLStr = SQLStr + " AND Status=0"
            End If
        Else
            SQLStr$ = "SELECT * FROM Kunden"
            For i% = 0 To (AnzSuchTeile% - 1)
                If (i% = 0) Then
                    SQLStr$ = SQLStr$ + " WHERE "
                Else
                    SQLStr$ = SQLStr$ + " AND "
                End If
                SQLStr$ = SQLStr$ + SuchName$(i%) + " " + SuchOp$(i%) + " '" + SuchWert$(i%) + "'"
            Next i%
            If (AlleKunden% = 0) Then
                SQLStr = SQLStr + " AND Status=0"
            End If
            SQLStr$ = SQLStr$ + " ORDER BY Name, Vorname"
        End If
        
        On Error Resume Next
        MatchTaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        MatchTaxeRec.Open SQLStr$, KundenDB1.ActiveConn

        If (MatchTaxeRec.EOF) Then
            SuchArtikel% = False
        Else
    '        Call Umspeichern("", 0)
    '        TaxeRec.MoveLast
    '        ListeSize% = TaxeRec.RecordCount
        End If
    End If
    
    If (SuchArtikel% = True) Then
        Call SuchArtikelTrue(IstPzn%)
    End If
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
'Dim row%
'Dim pzn$
'
'If (AlphaText$ <> "") Then
'    Call ListeBefuellen("9999999", AlphaText$, "", "")
'Else
'
'    With frmMatchcode.flxarbeit(0)
'        row% = .row
'
'        pzn$ = Format$(Ausgabe(row% - 1).pzn, "0000000")
'        Call ListeBefuellen(pzn$, .TextMatrix(row%, 1), .TextMatrix(row%, 2), .TextMatrix(row%, 3))
'    End With
'End If
'
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
'Dim i%, erg%, row%, ManuellAsatz%, ManuellSsatz%, ManuellBm%, ManuellNm%, abl%, ind%, iMenge%
'Dim h$, h2$, AufLager$
'
'iMenge% = 1
'If (Dp04Menge% > 0) Then iMenge% = Dp04Menge%
'
'If (pzn$ = "9999999") Then
'    h$ = "9999999" + vbTab + txt1$ + vbTab + vbTab + vbTab
'    h$ = h$ + Str$(iMenge%)
'Else
'    ManuellAsatz% = 0
'    FabsErrf% = hTax1.IndexSearch(0, pzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        ManuellAsatz% = CInt(FabsRecno&)
'        hTax1.GetRecord (FabsRecno& + 1)
'        iMenge% = hTax1.meng
'    End If
'
'    If (Trim(txt1$) = "") Then
'        txt1$ = hTax1.kurz
'        txt2$ = hTax1.meng
'        txt3$ = hTax1.meh
'    End If
'
'    h$ = pzn$ + vbTab + Trim$(txt1$) + vbTab + Trim$(txt2$) + vbTab + Trim$(txt3$)
'    h$ = h$ + vbTab + Str$(iMenge%)
'End If
'
'With frmMatchcode.flxarbeit(1)
'    .AddItem h$, 1
'    .row = 1
'    .col = 3
'    .col = 4
'    .ColSel = 4
'
'    .Visible = True
'    frmMatchcode.flxarbeit(0).Visible = False
'    frmMatchcode.txtFlexBack.Visible = False
'
'    frmMatchcode.txtMatchcode.SetFocus
'
'End With
'
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

Function CheckMaske%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckMaske%")
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
Dim i%, ret%, ind%
Dim sMaske$, AktMaske$, sWert$, AktWert$

ret% = True

sMaske$ = KundenMaske$
For i% = 0 To 3
    ind% = InStr(sMaske$, vbTab)
    If (ind% > 0) Then
        AktMaske$ = Trim(Left$(sMaske$, ind% - 1))
        sMaske$ = Mid$(sMaske$, ind% + 1)
        
        If (AktMaske$ <> "") Then
            Select Case i%
                Case 0
                    sWert$ = Trim(Kun1.kurz)
                Case 1
                    sWert$ = Trim(Kun1.Name(0)) + "," + Trim(Kun1.Name(1))
                Case 2
                    sWert$ = Trim(Kun1.Name(2)) + "," + Trim(Kun1.Name(3))
                Case 3
                    sWert$ = Trim(Kun1.Auswahl)
            End Select
            
            If (i% < 2) Then  'Komma in Kurz und Auswahl rausfiltern, sonst geht UND nicht
                Do
                    ind% = InStr(sWert$, ",")
                    If (ind% > 0) Then
                        sWert$ = Left$(sWert$, ind% - 1) + Mid$(sWert$, ind% + 1)
                    Else
                        Exit Do
                    End If
                Loop
            End If
            
            If (Right$(sWert$, 1) <> ",") Then
                sWert$ = sWert$ + ","
            End If
            
            Do
                ind% = InStr(sWert$, ",")
                If (ind% <= 0) Then
                    Exit Do
                End If
                
                AktWert$ = Left$(sWert$, ind% - 1)
                sWert$ = Mid$(sWert$, ind% + 1)
                
                ret% = clsOpTool.PatternTest(AktMaske$, AktWert$)
                If (ret%) Then
                    Exit For
                End If
            Loop
        End If
    Else
        Exit For
    End If
Next i%

CheckMaske% = ret%

Call clsError.DefErrPop
End Function

Function MakeMatchCode$(Name$)
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
Dim i%, iAsc%
Dim h$, ch$

h$ = Left$(Name$, 6)

For i% = 1 To Len(h$)
    ch$ = Mid$(h$, i%, 1)
    
    iAsc% = Asc(ch$)
    If (iAsc% < 32) Or (iAsc% = 92) Then
        ch$ = " "
    End If
    ch$ = UCase(ch$)
    
    Mid$(h$, i%, 1) = ch$
Next i%

MakeMatchCode$ = Left$(h$ + Space$(6), 6)

Call clsError.DefErrPop
End Function


Public Property Get InfoTyp%()
InfoTyp% = 5
End Property


Public Property Get WarteMaske%()
End Property
Public Property Let WarteMaske(ByVal vNewValue%)
End Property



