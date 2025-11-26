VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatchKunden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const INI_SECTION = "Auswahl-Kunden"
Const ARBEIT_SECTION = "Arbeitsbereich Auswahl-Kunden"
Const INFO_SECTION = "Infobereich Auswahl-Kunden"

Const EINGABE_NAME = "&Name/Nr"

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

Dim KundenMaske$

'Dim OrgSuchBegriffe$(10), OrgSb$(10)
'Dim AnzSuchBegriffe%


Private Const DefErrModul = "MATCHKUNDEN.CLS"

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
Else
    Ausgabe(pos%).pzn = Kun1.FabsRecno
    Ausgabe(pos%).Bookmark = Kun1.FabsRecno
    #If (FABSP = -1) Then
        Ausgabe(pos%).key = clsfabs.LastKey(6)
    #Else
        Ausgabe(pos%).key = String("A", 6)
    #End If

    h$ = " "
    h2$ = RTrim$(Kun1.kurz)
    h$ = h$ + vbTab + h2$ + vbTab + Mid$(Str$(Kun1.FabsRecno), 2)
    'Call OemToChar(h$, h$)
    
    Ausgabe(pos%).Name = h$
    Ausgabe(pos%).Verweis = Len(h$)
    Ausgabe(pos%).LagerKz = 0
End If

Call clsError.DefErrPop
End Sub

Function MachTaxeAuswahlName$(tRec As Recordset)
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
Else
    AltRecno& = Ausgabe(ind%).Bookmark
    AltKey$ = Left$(RTrim$(Ausgabe(ind%).key) + Space$(6), 6)
    
    Do
        If (Typ%) Then
            FabsErrf% = Kun1.IndexNext(0, AltRecno&, AltKey$, FabsRecno&)
        Else
            FabsErrf% = Kun1.IndexPrevious(0, AltRecno&, AltKey$, FabsRecno&)
        End If
        If (FabsErrf% = 0) Then
            Kun1.GetRecord (FabsRecno& + 1)
            If (CheckMaske%) Then
                SuchWeiter% = True
                Exit Do
            Else
                AltRecno& = FabsRecno&
                #If (FABSP = -1) Then
                    AltKey$ = clsfabs.LastKey(6)
                #Else
                    AltKey = String("A", 6)
                #End If
            End If
        Else
            Exit Do
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
        
        Set MatchTaxeRec = KundenDB.OpenRecordset(SQLStr$)

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
        Set MatchTaxeRec = KundenDB.OpenRecordset("Kunden", dbOpenTable)
        MatchTaxeRec.index = "Kurz"
        MatchTaxeRec.Seek ">=", SuchWert$(0)
        If (MatchTaxeRec.NoMatch = True) Then
            SuchArtikel% = False
        Else
    '        Call Umspeichern("", 0)
            If (AlleKunden% = 0) Then
                SuchArtikel% = False
                Do
                    If (clsOpTool.CheckNullLong(MatchTaxeRec!Status) = 0) Then
                        SuchArtikel = True
                        Exit Do
                    End If
                    MatchTaxeRec.MoveNext
                    If (MatchTaxeRec.EOF) Then
                        Exit Do
                    End If
                Loop
            End If
        End If
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
        
        Set MatchTaxeRec = KundenDB.OpenRecordset(SQLStr$)

        If (MatchTaxeRec.EOF) Then
            SuchArtikel% = False
        Else
    '        Call Umspeichern("", 0)
    '        TaxeRec.MoveLast
    '        ListeSize% = TaxeRec.RecordCount
        End If
    End If
    
'    If ((TextSuche% = True) Or (IstPzn% = True)) Then
'    '    Call StartAnimation(frmAction.aniSuchen)
'        If (IstPzn% = True) Then
'            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr = " + Such$
'        Else
'            If (Left$(Such$, 1) = "<") Then Such$ = Mid$(Such$, 2) + "*"
'            ind% = InStr(Such$, ",")
'            If (ind% > 0) Then
'                SQLStr$ = "SELECT * FROM Kunden WHERE Name LIKE '*" + Left$(Such$, ind% - 1) + _
'                "*' AND Name LIKE '*" + Mid$(Such$, ind% + 1) + "*'"
'            Else
'                ind% = InStr(Such$, ":")
'                If (ind% > 0) Then
'                    SQLStr$ = "SELECT * FROM Kunden WHERE Name LIKE '*" + Left$(Such$, ind% - 1) + _
'                    "*' OR Name LIKE '*" + Mid$(Such$, ind% + 1) + "*'"
'                Else
'                    ch$ = Right$(Such$, 1)
'                    If (ch$ <> "*") Then
'                        If (ch$ = ">") Then
'                            Such$ = "*" + Left$(Such$, Len(Such$) - 1)
'                        Else
'                            Such$ = Such$ + "*"
'                        End If
'                    End If
'                    SQLStr$ = "SELECT * FROM Kunden WHERE Name LIKE '" + Such$ + "'"
'                End If
'            End If
'        End If
'
'        Set MatchTaxeRec = KundenDB.OpenRecordset(SQLStr$)
'
'        If (MatchTaxeRec.EOF) Then
'            SuchArtikel% = False
'        Else
'    '        Call Umspeichern("", 0)
'    '        TaxeRec.MoveLast
'    '        ListeSize% = TaxeRec.RecordCount
'        End If
'    Else
'    End If
    
    If (SuchArtikel% = True) Then
        Call SuchArtikelTrue(IstPzn%)
    End If
Else
    If (InStr(Such$, vbTab) = 0) Then
        Such$ = Such$ + String(4, vbTab)
    End If
    
    KundenMaske$ = Such$
    
    If (Right$(Such$, 4) = String(4, vbTab)) Then
        nSuch$ = Left$(Such, Len(Such$) - 4)
        
        If (nSuch$ = "") Then nSuch$ = Space$(6)
        l% = Len(nSuch$)
        
        KundenMaske$ = ""
        For i% = 1 To l%
            ch$ = Mid$(nSuch$, i%, 1)
            If (InStr("!^*?><,:", ch$) > 0) Then
                KundenMaske$ = Such$
                Exit For
            End If
        Next i%
    
        If (l% > 6) Then
            l% = 6
            nSuch$ = Left$(nSuch$, l%)
        End If
        
        IstCode% = True
        For i% = 1 To l%
            ch$ = Mid$(nSuch$, i%, 1)
            If (InStr("0123456789", ch$) <= 0) Then
                IstCode% = False
                Exit For
            End If
        Next i%
        
        If (IstCode%) Then
            KundenMax& = (Kun1.DateiLen / Kun1.RecordLen)
            KundenNr& = Val(nSuch$)
            If (KundenNr& > KundenMax&) Then
                IstCode% = 0
            End If
        End If
        
        If (IstCode%) Then
            Kun1.GetRecord (KundenNr& + 1)
            
            If (iMatchNurEiner%) Then
                iMatchcodePzn$ = Format$(KundenNr&, "0000000")
                iMatchcodeErg$ = iMatchcodePzn$
                Unload frmMatchcode
            Else
                Call SuchArtikelTrue(True)
            End If
        
            Call clsError.DefErrPop: Exit Function
        End If
    End If
    
    If (KundenMaske$ = "") Then
    '    nSuch$ = such$
    Else
        nSuch$ = " "
    End If
        
    Call clsOpTool.InitPatternTest
    
    FabsErrf% = Kun1.IndexGeneric(0, nSuch$, FabsRecno&)
    If ((FabsErrf% = 0) Or (FabsErrf% = 12) Or (FabsErrf% = 13) Or (FabsErrf% = 15)) Then
        Do
            Kun1.GetRecord (FabsRecno& + 1)
            
            If (CheckMaske%) Then
                Call SuchArtikelTrue(False)
                Exit Do
            Else
                #If (FABSP = -1) Then
                    AltKey$ = clsfabs.LastKey(6)
                #Else
                    AltKey = String("A", 6)
                #End If
                FabsErrf% = Kun1.IndexNext(0, FabsRecno&, AltKey$, FabsRecno&)
            End If
            
            If (FabsErrf% <> 0) Then
                SuchArtikel% = False
                Exit Do
            End If
        Loop
    Else
        SuchArtikel% = False
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

'Function PatternTest%(sMaske$, sWert$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("PatternTest%")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim ok%, s1%, und%, oder%, x1%, x2%, xl%, ende%, found%
'Dim OrgText$, such$, Such1$
'
'If (sMaske$ = "") Then
'    PatternTest% = True: Call clsError.DefErrPop: Exit Function
'End If
'
'OrgText$ = sWert$
'
''Suchbegriff-Syntax:
''! irgendwo bedeutet "nicht"
''^ irgendwo bedeutet NICHT alles Upcase testen
''* Joker fÅr mehrere Zeichen (oder keines)
''? Joker fÅr ein Zeichen
''> letzer Suchbegriff mu· an letzter Stelle stehen
''< erster Suchbegriff mu· an erster Stelle stehen
'': ODER VerknÅpfung
'', UND VerknÅpfung
'
''Bsp.: !<^Asp*in
'' "!" Begriff gefunden -> Ergebnis = FALSE (nicht)
'' "^" NICHT alles Gro·buchstaben
'' "<" Begriff mu· am Anfang mit "Asp" beginnen
'' und danach irgendwo "in" haben
'
'such$ = sMaske$
'ok% = True
'
'Such1$ = such$
'While (Right$(Such1$, 1) = ",") Or (Right$(Such1$, 1) = ":")
'    Such1$ = Left$(Such1$, Len(Such1$) - 1)
'Wend
'
's1% = 1
'While (s1% > 0)
'    und% = 0
'    oder% = 0
'
'    If (Left$(Such1$, 1) = ",") Then
'        und% = True
'        Such1$ = Mid$(Such1$, 2)
'    End If
'    If (Left$(Such1$, 1) = ":") Then
'        oder% = True
'        Such1$ = Mid$(Such1$, 2)
'    End If
'
'    xl% = Len(Such1$)
'    x1% = InStr(Such1$, ":")
'    If (x1% = 0) Then
'        x1% = xl% + 1
'    End If
'    x2% = InStr(Such1$, ",")
'    If (x2% = 0) Then
'        x2% = xl% + 1
'    End If
'    If (x1% < x2%) Then
'        s1% = x1%
'    Else
'        s1% = x2%
'    End If
'
'    If (s1% > 0) And (s1% <= xl%) Then
'        such$ = Left$(Such1$, s1% - 1)
'        Such1$ = Mid$(Such1$, s1%)
'    Else
'        such$ = Such1$
'        Such1$ = ""
'        s1% = 0
'    End If
'
'    If (such$ <> "") Then
'        found% = PatternTest1(such$, OrgText$)
'    End If
'
'    ende% = 0
'    If (found% = 0) And (und%) Then
'        'nicht gefunden, war aber UND Kriterium
'        ok% = 0
'    ElseIf (found% <> 0) Then
'        If (und%) And Not (ok%) Then
'            ok% = 0
'        Else
'            ok% = True
'        End If
'    Else
'        'Bei "Oder" bleibt ok%, wie es war
'        If Not (oder%) Then
'            ok% = 0
'        End If
'    End If
'
'    If (ende%) Then
'        s1% = 0
'    End If
'Wend
'
'PatternTest% = ok%
'
'Call clsError.DefErrPop
'End Function
'
'Function PatternTest1%(such$, OrgText$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("PatternTest1%")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, ret%, ind%, ltext%, stern%, pos1%, Erst%, frage%, x1%, x2%, xl%, X%, p%
'Dim s$, Sonder$, sb$, text$, l$, such2$
'
'ret% = 0
'
'ind% = -1
'For i% = 1 To AnzSuchBegriffe%
'    If (OrgSuchBegriffe$(i% - 1) = such$) Then
'        ind% = i% - 1
'        Exit For
'    End If
'Next i%
'
'If (ind% >= 0) Then
'    sb$ = OrgSb$(ind%)
'ElseIf (AnzSuchBegriffe% < 10) Then
'    'Zerlegen:
'    s$ = such$
'
'    If (SuchSonder%("!", s$)) Then
'        sb$ = Chr$(1)
'    Else
'        sb$ = Chr$(0)
'    End If
'
'    If (SuchSonder%("^", s$)) Then
'        sb$ = sb$ + Chr$(0)
'    Else
'        sb$ = sb$ + Chr$(1)
'    End If
'
'    'Letzter Suchbegriff muﬂ ganz rechts stehen (nach trim)
'    If (SuchSonder%(">", s$)) Then
'      sb$ = sb$ + Chr$(1)
'    Else
'      sb$ = sb$ + Chr$(0)
'    End If
'
'    If (Asc(Mid$(sb$, 2, 1))) Then
'        s$ = UCase$(s$)
''        Call MyUCASE(s$)
'    End If
'    If (InStr(s$, " ") > 0) Then
'        s$ = RTrim(s$)
''        Call MyRTRIM(s$)
'    End If
'
'    sb$ = sb$ + s$
'
'    ind% = AnzSuchBegriffe%
'    OrgSuchBegriffe$(ind%) = such$
'    OrgSb$(ind%) = sb$
'
'    AnzSuchBegriffe% = AnzSuchBegriffe% + 1
'End If
'
'text$ = OrgText$
'If (InStr(text$, " ") > 0) Then
'    text$ = RTrim(text$)
''    Call MyRTRIM(text$)
'End If
'
'ltext% = Len(text$)
'If (Asc(Mid$(sb$, 2, 1))) Then
'    text$ = UCase$(text$)
''    Call MyUCASE(text$)
'End If
'
'such$ = Mid$(sb$, 4)
'If (Left$(such$, 1) = "<") Then
'    such$ = Mid$(such$, 2)
'ElseIf (such$ <> "") And (Left$(such$, 1) <> "*") And (Left$(such$, 1) <> "?") Then
'    such$ = "*" + such$
'End If
'
'ret% = 0
'
'stern% = 0
'pos1% = 0
'Erst% = -1
'
'While (such$ <> "")
'    l$ = Left$(such$, 1)
'    While (l$ = "*") Or (l$ = "?")
'        If (Left$(such$, 1) = "*") Then
'            stern% = -1
'            such$ = Mid$(such$, 2)
'        End If
'
'        frage% = 0
'        While (Left$(such$, 1) = "?")
'            frage% = frage% + 1
'            such$ = Mid$(such$, 2)
'        Wend
'        If (frage% > 0) Then
'            If (Len(text$) >= frage%) Then
'                text$ = Mid$(text$, frage% + 1)
'                frage% = 0
'            Else
'                ret% = 0
'                text$ = ""
'            End If
'        End If
'        l$ = Left$(such$, 1)
'    Wend
'
'    xl% = Len(such$)
'    x1% = InStr(such$, "?")
'    If (x1% = 0) Then
'        x1% = xl% + 1
'    End If
'    x2% = InStr(such$, "*")
'    If (x2% = 0) Then
'        x2% = xl% + 1
'    End If
'    If (x1% < x2%) Then
'        X% = x1%
'    Else
'        X% = x2%
'    End If
'
'    such2$ = Left$(such$, X% - 1)
'    such$ = Mid$(such$, X%)
'    p% = InStr(text$, such2$)
'    If (such$ = "") And (Asc(Mid$(sb$, 3)) > 0) And (p% > 0) Then
'        'letzter Begriff muﬂ an letzter Stelle stehen
'        p% = 0
'        If (Right$(text$, Len(such2$)) = such2$) Then
'            p% = Len(text$) - Len(such2$) + 1
'        End If
'    End If
'
'    If (text$ = "") And (such2$ = "") And (frage% = 0) Then
'      p% = 1
'    End If
'
'    If (p% = 0) Or (p% <> 1 And Not stern%) Then
'        ret% = 0
'        pos1% = 0
'        If (p% = 0) Or (Erst%) Then
'            such$ = ""
'        Else
'            such$ = Mid$(sb$, 4)
'            Erst% = -1
'        End If
'    Else
'        Erst% = 0
'        If (pos1% = 0) Then
'            pos1% = ltext% - Len(text$) + p%
'        End If
'        ret% = -1
'        text$ = Mid$(text$, p% + Len(such2$))
'    End If
'
'    stern% = 0
'Wend
'
'If (Asc(Mid$(sb$, 1, 1))) Then
'    ret% = Not (ret%)
'End If
'
'PatternTest1% = ret%
'
'Call clsError.DefErrPop
'End Function
'
'Function SuchSonder%(Sonder$, s$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("SuchSonder%")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim p%, X%, ret%
'
'ret% = True
'
'p% = 1
'While (p% > 0)
'    X% = InStr(p%, s$, Sonder$)
'    If (X% = 0) Then
'        ret% = 0
'        p% = 0
'    Else
'        If (X% > 1) Then
'            If (Mid$(s$, X% - 1, 1) = "\") Then
'                s$ = Left$(s$, X% - 2) + Mid$(s$, X%)
'                p% = X%
'                X% = 0
'            End If
'        End If
'        If (X% > 0) Then
'            s$ = Left$(s$, X% - 1) + Mid$(s$, X% + 1)
'            p% = 0
'        End If
'    End If
'Wend
'
'SuchSonder% = ret%
'
'Call clsError.DefErrPop
'End Function

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



