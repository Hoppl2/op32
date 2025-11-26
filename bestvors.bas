Attribute VB_Name = "modBestvors"
Option Explicit

Public BestvorsAbbruch%

Public BvProtAktiv%

Public ByteLeiste$
Dim AbholerPzns$

Dim bPzn$()
Dim nbPZN$()
Dim bkPZN$()
Dim besPZN$()
Dim bkPos%()
Dim bkLauf&()

Dim bMax%
Dim nbMax%
Dim bkMax%
Dim besMax%

Dim asatz%, ssatz%

Dim TaraMM%
Dim Retour%
Dim AbholerMenge%
Dim MinMen%, BestMen%

Dim xheute%

Dim AbholNummer%
Dim AnzArtikelDazu%

Dim DIREKTBEZUG%

Dim BvProt%

Dim LiefFuerHerst$(2)

Dim OhneLief%

Public DirektBevorratungsZeitraum%

Public DirektBezugsKz%

Public PalettenModus%

Public MitLagerstandCalc%, WuPruefung%

Private Const DefErrModul = "BESTVORS.BAS"

Sub BestellVorschlag(Optional HintergrundAktiv% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BestellVorschlag")
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
Dim i%, sMax%, VonSatz%, BisSatz%
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim lSatz&
Dim h$, SQLStr$

If (HintergrundAktiv%) Then
    BvProtAktiv% = False
ElseIf (ProgrammChar$ = "B") And (BvProtAktiv%) Then
    BvProtAktiv% = OpenBvProt%
Else
    BvProtAktiv% = False
End If

AnzArtikelDazu% = 0

h$ = Format(Now, "DDMMYY")
xheute% = iDate(h$)

'Call EinlesenWÜ
'
'If (para.BVRetour = "J") Then
'    Call EinlesenNB
'End If
'
'Call EinlesenBestellung
Call EinlesenWaMaArtikel

DIREKTBEZUG% = False
'DirektBezugsKz% = 0
'If (ProgrammChar$ = "B") And (IstDirektLief%) Then
If (IstDirektLief%) Then
    DIREKTBEZUG% = True
'    DirektBezugsKz% = 1
    lifzus.GetRecord (Lieferant% + 1)
    
    If (lifzus.ZuordnungenAktiv = 0) Then
        If (BvProtAktiv%) Then Close #BvProt%
        Call DefErrPop: Exit Sub
    End If
    
    DirektBevorratungsZeitraum% = lifzus.TempBevorratungsZeitraum
    If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = lifzus.BevorratungsZeitraum
    If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = para.BestellPeriode
    PalettenModus% = 0
    
    For i% = 0 To 2
        LiefFuerHerst$(i%) = UCase(Trim(lifzus.LiefFuerHerst(i%)))
    Next i%

    If (FremdPznOk%) Then
        SQLStr$ = "DELETE * FROM DIREKTAUFTEILUNG WHERE LIEF = " + Str$(Lieferant%)
        FremdPznDB.Execute (SQLStr$)
    End If
                

End If

If (DIREKTBEZUG% = False) Then
    If (InStr(para.Benutz, "Y") > 0) Then
        Call AbholerTesten
        AbholerMenge% = 0 'damit normale Bestellung nicht falsch wird
    End If
End If

Call ass.GetRecord(1)
'sMax% = ass.erstmax
sMax% = (ass.DateiLen / ass.RecordLen) - 1

If (sMax% <= 0) Then Call DefErrPop: Exit Sub

If (HintergrundAktiv%) Then

    VonSatz% = HintergrundSsatz%    '+ 1
    BisSatz% = VonSatz% + HintergrundAnz% - 1
'    If (VonSatz% = 1) Then VonSatz% = 2
    If (VonSatz% = 0) Then VonSatz% = 1
    
    For ssatz% = VonSatz% To BisSatz%
        If (ssatz% > sMax%) Then Exit For
        
        lSatz& = ssatz%
        Call ass.GetRecord(lSatz& + 1)
        
        If (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
            Call ArtikelInBestellung
        End If
    Next ssatz%
    
    HintergrundSsatz% = HintergrundSsatz% + HintergrundAnz%
    If (HintergrundSsatz% > sMax%) Then HintergrundSsatz% = 0
    
Else
    BestvorsAbbruch% = False

    frmBestVors!prgBestVors.Max = sMax%
    StartZeit! = Timer

    ass.GetRecord (1)

    For ssatz% = 1 To sMax%
        ass.GetRecord

        If (DIREKTBEZUG%) Then
'            If ((ass.lief = Lieferant%) Or (ass.lief = 0)) And (Val(ass.pzn) <> 0) Then Call ArtikelInPalette(ass.pzn)
            If (ass.Lief = Lieferant%) And (Val(ass.pzn) <> 0) Then Call ArtikelInPalette(ass.pzn)
        ElseIf (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
            Call ArtikelInBestellung
        End If

        If (ssatz% Mod 100 = 0) Then
            frmBestVors!lblBestVorsStatusWert(0).Caption = ssatz%
            frmBestVors!lblBestVorsStatusWert(1).Caption = AnzArtikelDazu%
            Dauer! = Timer - StartZeit!
            frmBestVors!lblBestVorsDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
            Prozent! = (ssatz% / sMax%) * 100!
            If (Prozent! > 0) Then
                GesamtDauer! = (Dauer! / Prozent!) * 100!
            Else
                GesamtDauer! = Dauer!
            End If
            RestDauer! = GesamtDauer! - Dauer!
            frmBestVors!lblBestVorsDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
            frmBestVors!prgBestVors.Value = ssatz%
            frmBestVors!lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
            
            h$ = Format$(Prozent!, "##0") + " %"
            With frmBestVors!picBestVorsProgress
                .Cls
                .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
                .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
                frmBestVors!picBestVorsProgress.Print h$
                frmBestVors!picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
'                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
            End With
            
            DoEvents
            If (BestvorsAbbruch% = True) Then
                Exit For
            End If
        End If
    Next ssatz%

    If (DIREKTBEZUG%) And (FremdPznOk%) And (IstAutoDirektLief% = 0) Then
        If (Trim(OpDirektPartner$) <> "") Then
            If (BestvorsAbbruch% = 0) Then
                Call FremdArtikelInPalette
            End If
        End If
    End If

End If

If (BvProtAktiv%) Then Close #BvProt%

Call DefErrPop
End Sub

Sub FremdArtikelInPalette()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FremdArtikelInPalette")
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
Dim i%, iProfilNr%, ls%, bMen%, Einfuegen%, pos%, vorhanden%, iMax%, rInd%
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!, bmo!
Dim lSatz&, sMax&
Dim h$, SQLStr$, pzn$
Dim FremdPznRec2 As Recordset

SQLStr$ = "SELECT DISTINCT pzn FROM Artikel WHERE LOKALAUFLAGER=0"
Set FremdPznRec2 = FremdPznDB.OpenRecordset(SQLStr$)
If (FremdPznRec2.EOF = False) Then
    FremdPznRec2.MoveLast
    sMax& = FremdPznRec2.RecordCount
End If

If (sMax& > 0) Then
    frmBestVors!prgBestVors.Max = sMax&
    StartZeit! = Timer
    
    FremdPznRec2.MoveFirst
    
    lSatz& = 0
    Do
        If (FremdPznRec2.EOF) Then
            Exit Do
        End If
        
        BestMen% = 0
        pzn$ = FremdPznRec2!pzn
        
        For i% = 1 To 2
            FremdPznRec.Seek "=", pzn$
            If (FremdPznRec.NoMatch = False) Then
                Do
                    If (FremdPznRec.EOF) Then
                        Exit Do
                    End If
                    If (FremdPznRec!pzn <> pzn$) Then
                        Exit Do
                    End If
                    
                    If (FremdPznRec!Lief = Lieferant%) Then
                        iProfilNr% = FremdPznRec!ProfilNr
                        If (iProfilNr% > 0) Then
                            h$ = Format(iProfilNr%, "000")
                            If (InStr(OpDirektPartner$, h$ + ",") > 0) Then
                                bmo! = FremdPznRec!opt
                                If (bmo! > 0) Then
                                    ls% = FremdPznRec!pos
                                    bMen% = CalcDirektBv%(bmo!, ls%)
                                    If (i% = 1) Then
                                        BestMen% = BestMen% + bMen%
                                    Else
                                        DirektAufteilungRec.AddNew
                                        DirektAufteilungRec!pzn = pzn$
                                        DirektAufteilungRec!ProfilNr = FremdPznRec!ProfilNr
                                        DirektAufteilungRec!Lief = FremdPznRec!Lief
                                        DirektAufteilungRec!bv = bMen%
                                        DirektAufteilungRec!BvGes = BestMen%
                                        DirektAufteilungRec.Update
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    FremdPznRec.MoveNext
                Loop
            End If
            If (i% = 1) Then
                If (BestMen% > 0) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    If (TaxeRec.EOF = False) Then
                        Call Taxe2ast(pzn$)
                    Else
                        BestMen% = 0
                    End If
                End If
                If (BestMen% = 0) Then
                    Exit For
                End If
            End If
        Next i%
        
        Einfuegen% = (BestMen% > 0)
        If (Einfuegen%) And (PalettenModus% = 0) Then
'            vorhanden% = VorhTest%(pos%)
'            If (vorhanden% = 1) Then
'                ww.SatzLock (1)
'                ww.GetRecord (1)
'                iMax% = ww.erstmax
'                rInd% = SucheDateiZeile(bkPos%(pos%), iMax%, bkLauf&(pos%))
'                If (rInd%) Then
'        '            If (ww.fixiert = Chr$(0)) And (ww.bm = ww.bm1) And (ww.nm = ww.nm1) And (ww.lief = ww.Lief1) And (ww.lief = Lieferant%) And (ww.auto = "+") Then
'                    If (ww.fixiert = Chr$(0)) And ((ww.Lief = Lieferant%) Or (ww.lief1 = Lieferant%)) And (ww.auto = "+") Then
'                        If (DirektBezugsKz% = 1) And (ww.DirektTyp = 0) Then
'                            ww.bm = BestMen%
'                            ww.PutRecord (rInd% + 1)
'                        Else
'                            If (ww.loesch = 0) Then
'                                ww.loesch = 1
'                                ww.status = 0
'                                ww.PutRecord (rInd% + 1)
'                            End If
'                            vorhanden% = 0
'                        End If
'                    End If
'                End If
'                ww.SatzUnLock (1)
'
'        '        ' neu
'        '        vorhanden% = 0
'            End If
'
'            If (vorhanden% <> 0) Then
'                Einfuegen% = False
'                'Anzeige warum nicht .....
'            End If
            
            If (Einfuegen%) Then
                asatz% = 0
                ssatz% = 1
                ass.Lief = Lieferant%
                Call Bestellen
            End If
        End If
        
        If (lSatz& Mod 100 = 0) Then
            frmBestVors!lblBestVorsStatusWert(0).Caption = lSatz&
            frmBestVors!lblBestVorsStatusWert(1).Caption = AnzArtikelDazu%
            Dauer! = Timer - StartZeit!
            frmBestVors!lblBestVorsDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
            Prozent! = (lSatz& / sMax&) * 100!
            If (Prozent! > 0) Then
                GesamtDauer! = (Dauer! / Prozent!) * 100!
            Else
                GesamtDauer! = Dauer!
            End If
            RestDauer! = GesamtDauer! - Dauer!
            frmBestVors!lblBestVorsDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
            frmBestVors!prgBestVors.Value = lSatz&
            frmBestVors!lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
            
            h$ = Format$(Prozent!, "##0") + " %"
            With frmBestVors!picBestVorsProgress
                .Cls
                .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
                .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
                frmBestVors!picBestVorsProgress.Print h$
                frmBestVors!picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
    '                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
            End With
            
            DoEvents
            If (BestvorsAbbruch% = True) Then
                Exit Do
            End If
        End If
        
        FremdPznRec2.MoveNext
        
        lSatz& = lSatz& + 1
    Loop
End If

Call DefErrPop
End Sub

Sub EinlesenWÜ()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWÜ")
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
Dim i%, j%, Max%

ww.GetRecord (1)
Max% = ww.erstmax
ReDim bPzn$(Max%)
j% = -1
For i% = 1 To Max%
    ww.GetRecord
    If (ww.status = 2) And (ww.IstAltLast = 0) And (ww.auto <> "v") Then
        j% = j% + 1
        bPzn$(j%) = ww.pzn + Chr$(ww.Lief)
'        bPZN$(j%) = Left$(wu.at, 7) + wu.li
    End If
Next i%
bMax% = j%

Call DefErrPop
End Sub

Sub EinlesenNB()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenNB")
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
Dim i%, j%, k%, Max%, gefunden%, gAlm%
Dim pzn$, gPzn$

ww.GetRecord (1)
Max% = ww.erstmax
ReDim nbPZN$(Max%)
j% = -1
For i% = 1 To Max%
    ww.GetRecord
    If (ww.status = 2) And (ww.IstAltLast) Then
        If (ww.WuNeuLm <> 0) Then
            gefunden% = False
            pzn$ = ww.pzn
            For k% = 0 To j%
                'Suchen, ob PZN schon da. Wenn ja - Menge addieren
                gPzn$ = Left$(nbPZN$(k%), 7)
                gAlm% = CVI(Mid$(nbPZN$(k%), 8, 2))
                If (gPzn$ = pzn$) Then
                    nbPZN$(k%) = gPzn$ + MKI(gAlm% + ww.WuNeuLm)
                    gefunden% = True
                    Exit For
                End If
            Next k%
            If (gefunden% = False) Then
                j% = j% + 1
                nbPZN$(j%) = pzn$ + MKI(ww.WuNeuLm)
            End If
        End If
    End If
Next i%
nbMax% = j%

Call DefErrPop
End Sub

Sub EinlesenWaMaArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWaMaArtikel")
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
Dim i%, j1%, j2%, j3%, k%, Max%, gAlm%, gefunden%
Dim pzn$, gPzn$

'ww.SatzLock (1)
Call ww.GetRecord(1)
Max% = ww.erstmax

ReDim bPzn$(Max%)
ReDim bkPZN$(Max%)
ReDim bkPos%(Max%)
ReDim bkLauf&(Max%)
ReDim nbPZN$(Max%)

j1% = -1
j2% = -1
j3% = -1

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    If (ww.status = 2) Then
        If (ww.IstAltLast = 0) And (ww.auto <> "v") Then
            j1% = j1% + 1
            bPzn$(j1%) = ww.pzn + Chr$(ww.Lief)
        ElseIf (ww.IstAltLast) And (ww.WuNeuLm <> 0) Then
            If (para.BVRetour = "J") Then
                gefunden% = False
                pzn$ = ww.pzn
                For k% = 0 To j3%
                    'Suchen, ob PZN schon da. Wenn ja - Menge addieren
                    gPzn$ = Left$(nbPZN$(k%), 7)
                    gAlm% = CVI(Mid$(nbPZN$(k%), 8, 2))
                    If (gPzn$ = pzn$) Then
                        nbPZN$(k%) = gPzn$ + MKI(gAlm% + ww.WuNeuLm)
                        gefunden% = True
                        Exit For
                    End If
                Next k%
                If (gefunden% = False) Then
                    j3% = j3% + 1
                    nbPZN$(j3%) = pzn$ + MKI(ww.WuNeuLm)
                End If
            End If
        End If
    ElseIf (ww.status = 1) And (ww.auto <> "v") Then
        j2% = j2% + 1
        bkPZN$(j2%) = ww.pzn + Chr$(ww.Lief)
        bkPos%(j2%) = i%
        bkLauf&(j2%) = ww.BekLaufNr
    End If
Next i%

bMax% = j1%
bkMax% = j2%
nbMax% = j3%
'ww.SatzUnLock (1)

Call DefErrPop
End Sub

Sub ArtikelInPalette(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ArtikelInPalette")
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
Dim i%, BestellMen%, Einfuegen%, vorhanden%, pos%, iMax%, rInd%, bMen%
Dim ls%, br%, BESUCH%, ok%, iProfilNr%
Dim bv!, bmo!, OrgBmo!
Dim iPzn$, ch$, h$

'If (pzn$ = "3487126") Then
'    Beep
'End If

If (PalettenModus% = 1) Then
    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ssatz% = CInt(FabsRecno&)
        ass.GetRecord (FabsRecno& + 1)
    Else
        ass.opt = 0
        ass.pp = " "
        ass.PosLag = 0
        ssatz% = 0
    End If
End If

If (MitLagerstandCalc%) Then
Else
    ass.PosLag = 0
End If

FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    asatz% = CInt(FabsRecno&)
    ast.GetRecord (FabsRecno& + 1)
    
'    If (PalettenModus% = 0) And (ass.lief = 0) Then
'        h$ = UCase(Trim(ast.herst))
'        ok% = False
'        If (h$ <> "") Then
'            For i% = 0 To 2
'                If (h$ = LiefFuerHerst$(i%)) Then
'                    ok% = True
'                    Exit For
'                End If
'            Next i%
'        End If
'        If (ok% = False) Then Call DefErrPop: Exit Sub
'        ass.lief = Lieferant%
'    End If
    
    'Optimierung:
    bmo! = ass.opt
    ls% = 0
    If (PosArtikel%()) Then
        ls% = ass.PosLag
    End If
    
    BestMen% = CalcDirektBv%(bmo!, ls%)
    
    'Partner-Apos
    If (FremdPznOk%) And (IstAutoDirektLief% = 0) Then
'    If (FremdPznOk%) Then
        For i% = 1 To 2
            FremdPznRec.Seek "=", pzn$
            If (FremdPznRec.NoMatch = False) Then
                Do
                    If (FremdPznRec.EOF) Then
                        Exit Do
                    End If
                    If (FremdPznRec!pzn <> pzn$) Then
                        Exit Do
                    End If
                    
                    If (FremdPznRec!Lief = Lieferant%) Then
                        iProfilNr% = FremdPznRec!ProfilNr
                        If (iProfilNr% > 0) Then
                            h$ = Format(iProfilNr%, "000")
                            If (InStr(OpDirektPartner$, h$ + ",") > 0) Then
                                bmo! = FremdPznRec!opt
                                If (bmo! > 0) Then
                                    ls% = FremdPznRec!pos
                                    bMen% = CalcDirektBv%(bmo!, ls%)
                                    If (i% = 1) Then
                                        BestMen% = BestMen% + bMen%
                                    Else
'                                        FremdPznRec.Edit
'                                        FremdPznRec!bv = bMen%
'                                        FremdPznRec!BvGes = BestMen%
'                                        FremdPznRec.Update
                                        DirektAufteilungRec.AddNew
                                        DirektAufteilungRec!pzn = pzn$
                                        DirektAufteilungRec!ProfilNr = FremdPznRec!ProfilNr
                                        DirektAufteilungRec!Lief = FremdPznRec!Lief
                                        DirektAufteilungRec!bv = bMen%
                                        DirektAufteilungRec!BvGes = BestMen%
                                        DirektAufteilungRec.Update
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    FremdPznRec.MoveNext
                Loop
            End If
        Next i%
    End If
        
    Einfuegen% = True
Else
    Einfuegen% = False
End If

If (Einfuegen% And (PalettenModus% = 0)) Then
    vorhanden% = VorhTest%(pos%)
    If (vorhanden% = 1) Then
        ww.SatzLock (1)
        ww.GetRecord (1)
        iMax% = ww.erstmax
        rInd% = SucheDateiZeile(bkPos%(pos%), iMax%, bkLauf&(pos%))
        If (rInd%) Then
'            If (ww.fixiert = Chr$(0)) And (ww.bm = ww.bm1) And (ww.nm = ww.nm1) And (ww.lief = ww.Lief1) And (ww.lief = Lieferant%) And (ww.auto = "+") Then
            If (ww.fixiert = Chr$(0)) And ((ww.Lief = Lieferant%) Or (ww.lief1 = Lieferant%)) And (ww.auto = "+") Then
                If (DirektBezugsKz% = 1) And (ww.DirektTyp = 0) Then
                    ww.bm = BestMen%
                    ww.PutRecord (rInd% + 1)
                Else
                    If (ww.loesch = 0) Then
                        ww.loesch = 1
                        ww.status = 0
                        ww.PutRecord (rInd% + 1)
                    End If
                    vorhanden% = 0
                End If
            End If
        End If
        ww.SatzUnLock (1)
        
'        ' neu
'        vorhanden% = 0
    End If
    
    If (vorhanden% <> 0) Then
        Einfuegen% = False
        'Anzeige warum nicht .....
    End If
End If
If (Einfuegen%) Then
    Call Bestellen
End If
    
Call DefErrPop
End Sub

Function CalcDirektBv%(opt!, ls%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcDirektBv%")
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
Dim br%, BESUCH%
Dim bmo!, bv!
    
'Optimierung:
If (opt! <= 0) Then
    bmo! = 0
Else
    bmo! = opt!
'        bmo% = Int(ass.opt + 0.501)
'        If (bmo% = 0) Then bmo% = 1
End If
    
    
'BvErrechnen:
br% = para.BestellPeriode
BESUCH% = DirektBevorratungsZeitraum%   ' para.BestellPeriode

bv! = (bmo! / br%) * BESUCH%
If (bv! < 0) Then bv! = 0

If (bv! > ls% + bmo!) Then
    bv! = bv! - ls%
Else
    bv! = bmo!
End If
If br% * ls% >= BESUCH% * (bmo! * 4 / 5) Then bv! = 0
If ls% = 0 And bv! = 0 Then bv! = 1
If bv! > 32000 Then bv! = 32000
If bv! < -32000 Then bv! = -32000

CalcDirektBv% = Int(bv! + 0.501)
    
Call DefErrPop
End Function

Sub ArtikelInBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ArtikelInBestellung")
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
Dim BestellMen%, Einfuegen%, vorhanden%, pos%, iMax%, rInd%, bMen%
Dim iPzn$, ch$

If ((Val(ass.pzn) = 0) Or (ass.halt = "S")) Then Call DefErrPop: Exit Sub

If (ass.vmm = 8224) Then
    ass.vmm = 0
End If

MinMen% = ass.MM
If ((ass.vmm > 0) Or ((ass.vmm = 0) And (ass.vbm > 0))) Then MinMen% = ass.vmm

Retour% = RetourenSuchen%

BestellMen% = ass.bm
If (ass.vbm > 0) Then BestellMen% = ass.vbm

'IF debug% THEN PRINT #f.deb%, USING "& ### ### ### ### ### ### ###";ass.pzn$; CVI(ass.PosLag$); CVI(ass.BM$); CVI(ass.vBM$); CVI(ass.MM$); CVI(ass.vMM$); Abholermenge%; Retour%;
If (BvProtAktiv%) Then Print #BvProt%, ass.pzn + iFormat(ass.PosLag, 4) + iFormat(ass.bm, 4) + iFormat(ass.vbm, 4) + iFormat(ass.MM, 4) + iFormat(ass.vmm, 4) + iFormat(AbholerMenge%, 4) + iFormat(Retour%, 4);

If (para.LagerAuffuellen) Then
    BestMen% = BestellMen% + AbholerMenge% + MinMen% - ass.PosLag
    If (para.BVRetour = "J") Then
        BestMen% = BestMen% - Retour%
    End If
Else
    BestMen% = BestellMen%
End If
'IF debug% THEN PRINT #f.deb%, USING " ###"; BestMen%;
If (BvProtAktiv%) Then Print #BvProt%, iFormat(BestMen%, 4);

If (para.TaraKontrolle) Then Call TaraKontrolleAction
'IF debug% THEN PRINT #f.deb%, USING " ###"; TaraMM%;
If (BvProtAktiv%) Then Print #BvProt%, iFormat(TaraMM%, 4);

Einfuegen% = TaraMMKontrollieren%
If (BvProtAktiv%) Then
'    Print #f.deb%, USING; " ### ##"; BestMen%; Bestellen%;
    Print #BvProt%, iFormat(BestMen%, 4) + iFormat(Einfuegen%, 3);
    If (ass.halt = "S") Then
        Print #BvProt%, " S";
    Else
        Print #BvProt%, "  ";
    End If
End If

If (Einfuegen%) Then
    'Test ob reduzieren und ob dann MM reicht
    Call ReduTest
'    IF debug% THEN PRINT #f.deb%, USING " ### ### ###"; CVI(ass.BM$); CVI(ass.MM$); BestMen%;
    If (BvProtAktiv%) Then Print #BvProt%, iFormat(ass.bm, 4) + iFormat(ass.MM, 4) + iFormat(BestMen%, 4);
    Einfuegen% = TaraMMKontrollieren%
'    IF debug% THEN PRINT #f.deb%, USING " ### ##"; BestMen%; Bestellen%;
    If (BvProtAktiv%) Then Print #BvProt%, iFormat(BestMen%, 4) + iFormat(Einfuegen%, 3);
End If
If (Einfuegen%) Then
    FabsErrf% = ast.IndexSearch(0, ass.pzn, FabsRecno&)
    If (FabsErrf% = 0) Then
        asatz% = CInt(FabsRecno&)
        ast.GetRecord (FabsRecno& + 1)
        Einfuegen% = PosArtikel%()
    Else
        Einfuegen% = False
    End If
    'IF debug% THEN PRINT #f.deb%, USING " ## ##"; errf%; PP%;
    If (BvProtAktiv%) Then Print #BvProt%, iFormat(FabsErrf%, 3) + iFormat(Einfuegen%, 3);
End If
If (Einfuegen%) Then
    vorhanden% = VorhTest%(pos%)
    If (vorhanden% = 1) Then
        ww.SatzLock (1)
        ww.GetRecord (1)
        iMax% = ww.erstmax
        rInd% = SucheDateiZeile(bkPos%(pos%), iMax%, bkLauf&(pos%))
        If (rInd%) Then
            If (ww.fixiert = Chr$(0)) And (ww.bm = ww.bm1) And (ww.nm = ww.nm1) And (ww.Lief = ww.lief1) Then
                If (BestMen% > 0) Then
                    bMen% = BestMen%
                Else
                    bMen% = 1
                End If
                
                If (ww.bm <> bMen%) Then
                    If (ww.loesch = 0) Then
                        ww.loesch = 1
                        ww.status = 0
                        ww.PutRecord (rInd% + 1)
                    End If
                    vorhanden% = 0
                End If
            End If
        End If
        ww.SatzUnLock (1)
    End If
    
    If (vorhanden% <> 0) Then
        Einfuegen% = False
        'Anzeige warum nicht .....
    End If
End If
If (Einfuegen%) Then
    If (ast.lic <> "?") Then
        Call Bestellen
    Else
'        IF debug% THEN PRINT #f.deb%, USING " & &"; ast.vc$; ast.lic$;
        If (BvProtAktiv%) Then Print #BvProt%, ast.vc; ast.lic;
    End If
End If

If (BvProtAktiv%) Then Print #BvProt%,
    
Call DefErrPop
End Sub

Sub TaraKontrolleAction()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaraKontrolleAction")
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
Dim i%, KleinstTaraPlatz%, iTaraPlatz%

TaraMM% = 0
If (ass.Max(1) = 0) Then Call DefErrPop: Exit Sub  'nur max. 1 Taralager
'If (ass.wert("max", 1) = 0) Then Call DefErrPop: Exit Sub  'nur max. 1 Taralager

'Tara-Mindestmenge = Summe der Plätze minus dem kleinsten plus 1
'damit kann nicht passieren, daß ein Platz leer ist und nix zum nachräumen da ist,
'selbst wenn alle anderen Lagerplätze voll sind und nur einer leer, wird bestellt

KleinstTaraPlatz% = 9999
For i% = 0 To 3
    iTaraPlatz% = ass.Max(i%)
    If (iTaraPlatz% > 0) Then
        TaraMM% = TaraMM% + iTaraPlatz%
        If (iTaraPlatz% < KleinstTaraPlatz%) Then
            KleinstTaraPlatz% = iTaraPlatz%
        End If
    End If
Next i%
TaraMM% = TaraMM% - KleinstTaraPlatz% + 1

If (TaraMM% <= MinMen%) Then TaraMM% = 0

Call DefErrPop
End Sub

'Sub TaraKontrolleAction()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TaraKontrolleAction")
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
'Dim i%, KleinstTaraPlatz%, iTaraPlatz%
'
'TaraMM% = 0
'If (ass.teinzel(1).teplatz = 0) Then Call DefErrPop: Exit Sub   'nur max. 1 Taralager
'
''Tara-Mindestmenge = Summe der Plätze minus dem kleinsten plus 1
''damit kann nicht passieren, daß ein Platz leer ist und nix zum nachräumen da ist,
''selbst wenn alle anderen Lagerplätze voll sind und nur einer leer, wird bestellt
'
'KleinstTaraPlatz% = 9999
'For i% = 0 To 3
'    iTaraPlatz% = ass.teinzel(i%).teplatz
'    If (iTaraPlatz% > 0) Then
'        TaraMM% = TaraMM% + iTaraPlatz%
'        If (iTaraPlatz% < KleinstTaraPlatz%) Then
'            KleinstTaraPlatz% = iTaraPlatz%
'        End If
'    End If
'Next i%
'TaraMM% = TaraMM% - KleinstTaraPlatz% + 1
'
'If (TaraMM% <= MinMen%) Then TaraMM% = 0
'
'Call DefErrPop
'End Sub

Function TaraMMKontrollieren%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaraMMKontrollieren%")
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
Dim assMM%, ret%

ret% = True

assMM% = MinMen%
If ((ass.PosLag - AbholerMenge% + Retour%) > MinMen%) Then
    If (TaraMM% > 0) Then
        MinMen% = TaraMM%
        BestMen% = ass.bm + Retour% + MinMen% - AbholerMenge% - ass.PosLag
        If ((ass.PosLag - AbholerMenge% + Retour%) > MinMen%) Then ret% = False
    Else
        ret% = False
    End If
    If ((InStr(para.Benutz, "*") > 0) And (ass.PosLag <= ass.flager) And (para.LstGemeinsam = "J")) Then
        BestMen% = assMM%
    End If
End If

TaraMMKontrollieren% = ret%

Call DefErrPop
End Function

Function RetourenSuchen%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RetourenSuchen%")
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
Dim i%, ret%

ret% = 0
If (para.BVRetour = "J") Then
    For i% = 0 To nbMax%
        If (ass.pzn = Left$(nbPZN$(i%), 7)) Then
            ret% = CVI(Mid$(nbPZN$(i%), 8, 2))
            Exit For
        End If
    Next i%
End If

RetourenSuchen% = ret%

Call DefErrPop
End Function

Sub ReduTest()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReduTest")
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
Dim BmStamm%, MmStamm%, verbrauch%, zeit%, lld%, llm%
Dim BmoAkt!
Dim xc$

BmStamm% = ass.bm
MmStamm% = ass.MM
If ((BmStamm% = 1) And (MmStamm% = 0)) Then Call DefErrPop: Exit Sub

verbrauch% = POSverbrauch%(zeit%)

If (verbrauch% < 1) Then
    lld% = ass.lld
    llm% = ass.llm
    verbrauch% = ass.lag + llm% - ass.PosLag
    zeit% = xheute% - lld%
End If
If (zeit% < 1) Then Call DefErrPop: Exit Sub

BmoAkt! = CSng(verbrauch%) / CSng(zeit%) * para.bp1
If (BmoAkt! < Abs(BmStamm%)) Then Call Reduzieren

Call DefErrPop
End Sub

Sub Reduzieren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Reduzieren")
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
Call DefErrPop
End Sub

'Function POSverbrauch%(zeit%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("POSverbrauch%")
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
'Dim j%, X%, verbrauch%
'Dim xc$
'
'verbrauch% = 0
'For j% = 9 To 0 Step -1
'    xc$ = ass.verk(j%).Datum
'    X% = CVDat%(xc$)
'    If ((X% + para.bp1 * 2) >= xheute%) Then
'        verbrauch% = verbrauch% + ass.verk(j%).Menge
'        zeit% = xheute% - X%
'    End If
'Next j%
'
'POSverbrauch% = verbrauch%
'
'Call DefErrPop
'End Function

Function POSverbrauch%(zeit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("POSverbrauch%")
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
Dim j%, X%, verbrauch%
Dim xc$

verbrauch% = 0
For j% = 9 To 0 Step -1
    X% = ass.vkdat(j%)
    If ((X% + para.bp1 * 2) >= xheute%) Then
        verbrauch% = verbrauch% + ass.vkmng(j%)
        zeit% = xheute% - X%
    End If
Next j%

POSverbrauch% = verbrauch%

Call DefErrPop
End Function

Function PosArtikel%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PosArtikel%")
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
Dim ret%

ret% = False
If (InStr("AF", ass.pp) > 0) Then ret% = True
If (InStr(para.PosAktivWG, Left$(ast.wg, 1)) > 0) Then ret% = True

PosArtikel% = ret%

Call DefErrPop
End Function

Function VorhTest%(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("VorhTest%")
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
Dim i%, mm1%, gefunden%, AnzInBest%, AnzInWü%, AnzFuerGh%, DirLief%, ret%, iBestMen%
Dim nUeber%, OrgRet%
Dim SQLStr$, h$

OhneLief% = False

'schon in der Bestellung?
ret% = 0
AnzInBest% = 0
AnzFuerGh% = 0
AnzInWü% = 0
For i% = 0 To bkMax%
    If (Left$(bkPZN$(i%), 7) = ass.pzn) Then
        ret% = 1
        AnzInBest% = AnzInBest% + 1
        
        pos% = i%
        DirLief% = Asc(Mid$(bkPZN$(i%), 8, 1))
        If (DirLief% = 0) Then
            AnzFuerGh% = AnzFuerGh% + 1
        Else
            h$ = Format(DirLief%, "000")
            If (InStr(OpPartnerLiefs$, h$ + ",") > 0) Then
                AnzFuerGh% = AnzFuerGh% + 1
            End If
        End If
    End If
Next i%
If (ret% = 1) Then
    If (DIREKTBEZUG%) Then
    ElseIf (AnzInBest% > 1) And (AnzFuerGh% > 0) Then
        ret% = 3
    Else
        AnzInWü% = 1
    End If
End If

'WÜ auch prüfen ???
If (ret% = 0) Then
    If (WuPruefung%) Or (DIREKTBEZUG% = False) Then
        'man muß wissen, ob Artikel mehrmals in WÜ steht
        AnzInWü% = 0
        DirLief% = 0
        For i% = 0 To bMax%
            If (Left$(bPzn$(i%), 7) = ass.pzn) Then
                ret% = 2
                AnzInWü% = AnzInWü% + 1
                DirLief% = Asc(Mid$(bPzn$(i%), 8, 1))
                If (AnzInWü% > 1) Then Exit For
            End If
        Next i%
    End If
End If

'IF debug% THEN PRINT #f.deb%, USING " # # ###"; vorh%; AnzInWue%; DirLief%;
If (BvProtAktiv%) Then Print #BvProt%, iFormat(ret%, 2) + iFormat(AnzInWü%, 2) + iFormat(DirLief%, 4);

'If ((ret% = 2) And (AnzInWü% = 1) And (para.DirektLieferanten <> "")) Then
If ((AnzInWü% = 1) And (para.DirektLieferanten <> "")) Then
    'bei Artikeln mit MM>1 prüfen, ob Lagerstand<=MM/2. Wenn ja, auf MM auffüllen
    '(Defekt bei Direktlieferantenbestellung verhindern)
    Call DirektLieferant(DirLief%, para.DirektLieferanten, gefunden%)
    If (gefunden%) Then
        OrgRet% = ret%
        mm1% = ass.vmm
        If (mm1% <= 0) Then mm1% = ass.MM
        If ((mm1% >= 1) And (ass.PosLag <= Int(mm1% / 2))) Then
            ret% = 0
            iBestMen% = mm1% - ass.PosLag
        End If
        If (mm1% = 0) And (ass.PosLag = 0) Then
            ret% = 0
            iBestMen% = 1
        End If
        If (ret% = 0) Then
            nUeber% = 0
            'nur wenn über GH beziehbar !
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ass.pzn
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                If (TaxeRec!UeberGH = 1) Then
                    nUeber% = True
                End If
            End If
            If (nUeber%) Then
                ret% = OrgRet%  '2
            Else
                BestMen% = iBestMen%
                OhneLief% = True
            End If
        End If
    End If
End If
'IF debug% THEN PRINT #f.deb%, USING " ###"; BestMen%;
If (BvProtAktiv%) Then Print #BvProt%, iFormat(BestMen%, 4);

VorhTest% = ret%

Call DefErrPop
End Function

Sub DirektLieferant(AktLief%, DirLief$, gefunden%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektLieferant")
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
Dim s%, von%, bis%
Dim Lief$, Such$

'Suche, ob der WÜ-Lieferant unter den in PDATEI angeführten Direktlieferanten ist
gefunden% = False
Lief$ = DirLief$
While ((Lief$ <> "") And Not (gefunden%))
    s% = InStr(Lief$, ",")
    If (s% > 0) Then
        Such$ = Left$(Lief$, s% - 1)
        Lief$ = Mid$(Lief$, s% + 1)
    Else
        Such$ = Lief$
        Lief$ = ""
    End If
    
    s% = InStr(Such$, "-")
    If (s% > 0) Then
        von% = Val(Left$(Such$, s% - 1))
        bis% = Val(Mid$(Such$, s% + 1))
        If (bis% = 0) Then bis% = 200
        If ((von% <= AktLief%) And (AktLief% <= bis%)) Then gefunden% = True
    Else
        If (Val(Such$) = AktLief%) Then gefunden% = True
    End If
Wend

Call DefErrPop
End Sub

Sub Bestellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Bestellen")
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
Dim i%, X%, Max%
Dim xc$, iSatz$

ww.pzn = ast.pzn
ww.txt = ast.kurz + ast.meng + ast.meh
ww.asatz = asatz%
ww.ssatz = ssatz% '??
ww.nm = 0
ww.aep = ast.aep
ww.AVP = ast.AVP
ww.abl = ast.abl: If (ww.abl <> "A") Then ww.abl = "X"
ww.wg = Left$(ast.wg, 1)
ww.best = " "
ww.absage = 0
ww.bm = 1
ww.km = 1
ww.angebot = 0
ww.Lief = 0
ww.auto = "+"
ww.alt = Chr$(0)
ww.nnart = 0
ww.NNAEP = 0#

ww.zugeordnet = Chr$(0)
ww.zukontrollieren = Chr$(0)
ww.fixiert = Chr$(0)
'bek.musskontrollieren = Chr$(0)
ww.DirektTyp = DirektBezugsKz% ' Chr$(0) wegen Direktbezug Einlesen

For i% = 0 To 5
    ww.actkontrolle(i%) = 111
Next i%
ww.actzuordnung = 111
ww.aktivlief = 0
ww.aktivind = 0
ww.PosLag = 0

ww.herst = ast.herst

'If (ast.Rez = "SG") Then
'    ww.IstBtm = 1
'Else
'    ww.IstBtm = 0
'End If

ww.zr = -123.45!
                            
ww.loesch = 0
ww.status = 1

ww.ArtikelKz = 0
ww.ArtikelKz2 = 0

If (ssatz% <> 0) Then
    ww.Lief = ass.Lief
    If (BestMen% > 0) Or (DIREKTBEZUG%) Then
        ww.bm = BestMen%
        ww.km = ass.bm
    End If
    
    X% = ass.lld
    If ((X% + para.MonNBest) < xheute%) Then ww.alt = "?"
End If

If (PalettenModus% = 1) Then ww.Lief = Lieferant%
If (OhneLief%) Then
    ww.Lief = 0
    ww.ArtikelKz2 = KZ_ISTINTERIM
End If
OhneLief% = False

ww.BekLaufNr = 0

ww.IstSchwellArtikel = 0
ww.OrgZeit = 0

ww.SatzLock (1)
ww.GetRecord (1)
bkMax% = ww.erstmax
bkMax% = bkMax% + 1
ww.erstmax = bkMax%
ww.PutRecord (1)
ww.PutRecord (bkMax% + 1)
ww.SatzUnLock (1)

ReDim Preserve bkPZN$(bkMax%)
bkPZN$(bkMax%) = ast.pzn + Chr$(ww.Lief)

AnzArtikelDazu% = AnzArtikelDazu% + 1

BekartCounter% = -1

Call DefErrPop
End Sub

Sub AbholerTesten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerTesten")
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
Dim i%, j%, erg%, KistenNr%, AbholerStatus%, gefunden%, gMenge%
Dim h$, pzn$, gPzn$

kiste.OpenDatei

besMax% = -1
For KistenNr% = 1 To 999
    If (kiste.Belegt(KistenNr%)) Then
        kiste.GetKiste (KistenNr%)
        
        For i% = 0 To 9
            erg% = kiste.GetInhalt(i%)
            If (erg%) Then
                If (kiste.WasTun = "B") Then
                    AbholerStatus% = kiste.status
                    If (AbholerStatus% = 1) Then
                        'fertig, aber nicht abgeholt
                        AbholerMenge% = 1
                        For j% = 1 To 10
                            h$ = RTrim$(kiste.InfoText(j% - 1))
                            If Mid$(h$, 18, 3) = " x " Then
                                AbholerMenge% = -Val(Mid$(h$, 21, 4))
                            End If
                            If Mid$(h$, 14, 4) = "PZN=" Then
                                pzn$ = Mid$(h$, 18, 7)
                            End If
                        Next j%
                        If ((Val(pzn$) > 0) And (pzn$ <> "999999")) Then
                            gefunden% = False
                            For j% = 0 To besMax%
                                'Suchen, ob PZN schon da. Wenn ja - Menge addieren
                                gPzn$ = Left$(besPZN$(j%), 7)
                                gMenge% = CVI(Mid$(besPZN$(j%), 8, 2))
                                If (gPzn$ = pzn$) Then
                                    besPZN$(j%) = gPzn$ + MKI(gMenge% + AbholerMenge%)
                                    gefunden% = True
                                    Exit For
                                End If
                            Next j%
                            If (gefunden% = False) Then
                                besMax% = besMax% + 1
                                ReDim Preserve besPZN$(besMax%)
                                besPZN$(besMax%) = pzn$ + MKI(AbholerMenge%)
                            End If
                        End If
                    End If
                End If
            End If
        Next i%
    End If
Next KistenNr%

kiste.CloseDatei

AbholerPzns$ = ""
For i% = 0 To besMax%
    pzn$ = Left$(besPZN$(i%), 7)
    AbholerMenge% = CVI(Mid$(besPZN$(i%), 8, 2))
    If (AbholerMenge% <> 0) Then
        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ssatz% = FabsRecno&
            Call ass.GetRecord(FabsRecno& + 1)
            Call ArtikelInBestellung
        End If
        AbholerPzns$ = AbholerPzns$ + " " + pzn$
    End If
Next i%

Call DefErrPop
End Sub

Function OpenBvProt%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenBvProt%")
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
Dim ret%
Dim t$

ret% = False

On Error Resume Next
BvProt% = FreeFile
Open "BVDEB.$$$" For Output Access Write Shared As #BvProt%

If (Err = 0) Then
    ret% = True
    
    Print #BvProt%, "POS-WG: "; para.PosAktivWG; " DirLief: "; para.DirektLieferanten;
    Print #BvProt%, " Tara: ";
    If (para.TaraKontrolle) Then Print #BvProt%, "J"; Else Print #BvProt%, "N";
    Print #BvProt%, " Nachbe: "; para.BVRetour;
    'If land$ <> "D" Then
    '    Print #BvProt%, " Auffüllen: ";
    '    If (para.LagerAuffuellen) Then Print #BvProt%, "J"; Else Print #BvProt%, "N";
    'End If
    Print #BvProt%, " Abholer: ";
    If InStr(para.Benutz, "Y") > 0 Then Print #BvProt%, "J"; Else Print #BvProt%, "N";
    Print #BvProt%,
    t$ = "        Lag  BM vBM  MM vMM Abh  NB  B1  VK  B2 ok S BMn MMn  B3  B4 ok ER PP ? W Lif  B5 V L"
    Print #BvProt%, ".#5.KOPF"; t$
    Print #BvProt%, ".KOPF"; String$(Len(t$), "-")
Else
    BvProt% = 0
End If

OpenBvProt% = ret%

Call DefErrPop
End Function

Function iFormat$(Wert%, anz%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iFormat$")
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
iFormat$ = Right$(Space$(anz%) + Str$(Wert%), anz%)
Call DefErrPop
End Function


