VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialoge"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DefErrModul = "DIALOGE.CLS"

'Function MatchCode$(mTyp%, pzn$, txt$, Optional AutoRet% = False, Optional NurEiner% = False)
Function Matchcode$(mTyp%, pzn$, txt$, AutoRet%, NurEiner%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MatchCode$")
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
Dim ret%, FremdPznFlag%
Dim h$

ret% = False
FremdPznFlag% = 0
SonderPrOk% = 0

Dp04Ok% = False
SeriellScannerOk% = False

If (mTyp% = 1) Then
'    Set Match2 = New clsMatchLieferanten
'    Match2.MatchTyp = MATCH_LIEFERANTEN
'
'    Match2.MatchAutoRet = AutoRet%
'    Match2.MatchNurEiner = True
'
'    Match2.MatchcodePzn = "0000000"
'    Match2.MatchcodeTxt = Trim(txt$)
'    Match2.MatchcodeErg = ""
'    Call Match2.InitMatch
'    frmLiefMatchcode.Show 1
'    pzn$ = Match2.MatchcodePzn
'    txt$ = Match2.MatchcodeTxt
'
'    Matchcode$ = Match2.MatchcodeErg
'
'    Set Match2 = Nothing
    Set Match1 = New clsMatchLieferanten
    Match1.MatchTyp = MATCH_LIEFERANTEN

    Match1.MatchAutoRet = AutoRet%
    Match1.MatchNurEiner = True
    
    Match1.MatchcodePzn = clsOpTool.PznString(0)
    Match1.MatchcodeTxt = Trim(txt$)
    Match1.MatchcodeErg = ""
    Call Match1.InitMatch
    frmMatchcode.Show 1
    pzn$ = Match1.MatchcodePzn
    txt$ = Match1.MatchcodeTxt
    
    Matchcode$ = Match1.MatchcodeErg
    
    Set Match1 = Nothing
    
    If (LieferantenDBok) Then
        On Error Resume Next
        LieferantenRec.Close
        Err.Clear
        On Error GoTo DefErr
    End If
Else
    If (mTyp% = 0) Or (mTyp% = 2) Or (mTyp% = 4) Then
        If (TaxeAdoDBok) Then
            Set Match1 = New clsMatchArtikelADO
        Else
            Set Match1 = New clsMatchArtikel
        End If
        Match1.MatchTyp = MATCH_ARTIKEL
        If (mTyp% = 0) Then
            Match1.MatchListeTyp = 0
        ElseIf (mTyp% = 2) Then
            Match1.MatchListeTyp = 1
        Else
            Match1.MatchListeTyp = 2
        End If
        Dp04Ok% = clsOpTool.ModemParameter("DP-0", h$, False)
        SeriellScannerOk% = clsOpTool.ModemParameter("SCANNER", SeriellScannerParam$, False)
        
        If (FremdPznTest% = 0) Then
            FremdPznOk% = OpenOpConnect(FremdPznDB, OpPartnerDB)
            FremdPznTest% = 0
            FremdPznFlag% = True
        End If
        
        If (ArtikelDBok) Then
        Else
            On Error Resume Next
            Err.Clear
            Set SonderPrDB = OpenDatabase(SONDERPR_DB, False, True)
            SonderPrOk% = (Err.Number = 0)
            Err.Clear
            On Error GoTo DefErr
        End If
    Else
        If (TaxeAdoDBok) Then
            Set Match1 = New clsMatchHilfsTaxeADO
        Else
            Set Match1 = New clsMatchHilfsTaxe
        End If
        Match1.MatchTyp = MATCH_HILFSTAXE
        FremdPznOk% = 0
    End If

    Match1.MatchAutoRet = AutoRet%
            
    Match1.MatchNurEiner = NurEiner%
    If (mTyp% = 3) Then
        Match1.MatchNurEiner = True
    End If
    
    Match1.MatchcodePzn = clsOpTool.PznString(0)
    Match1.MatchcodeTxt = Trim(txt$)
    Match1.MatchcodeErg = ""
    Call Match1.InitMatch
    frmMatchcode.Show 1
    pzn$ = Match1.MatchcodePzn
    txt$ = Match1.MatchcodeTxt
    
    Matchcode$ = Match1.MatchcodeErg
    
    Set Match1 = Nothing
    
    If (FremdPznFlag%) And (FremdPznOk%) Then
        FremdPznDB.Close
        FremdPznOk% = 0
    End If
    
    If (SonderPrOk) Then
        SonderPrDB.Close
    End If
End If

MatchTyp% = -1

Call clsError.DefErrPop
End Function

'Function MatchcodeKunden$(mTyp%, pzn$, txt$, AutoRet%, NurEiner%, Optional StammProg% = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("MatchCodeKunden$")
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
'Dim ret%, FremdPznFlag%
'Dim h$
'Dim KundenDB1 As clsKundenDB
'
'ret% = False
'FremdPznFlag% = 0
'SonderPrOk% = 0
'
'Dp04Ok% = False
'SeriellScannerOk% = False
'
'MatchAusStammProg% = StammProg%
'
'Set KundenDB1 = New clsKundenDB
'KundenDBok = KundenDB1.DBvorhanden
'If (KundenDBok) Then
'    KundenDBok = KundenDB1.OpenDB
'End If
'
'If (KundenDBok) Then
'    Set Match1 = New clsMatchKundenADO
'Else
'    Set Match1 = New clsMatchKunden
'End If
'Match1.MatchTyp = MATCH_KUNDEN
'
'Match1.MatchAutoRet = AutoRet%
'Match1.MatchNurEiner = True ' NurEiner%
'
'Match1.MatchcodePzn = clsOpTool.PznString(0)
'Match1.MatchcodeTxt = Trim(txt$)
'Match1.MatchcodeErg = ""
'Call Match1.InitMatch
'frmMatchcodeKunden.Show 1
'pzn$ = Match1.MatchcodePzn
'txt$ = Match1.MatchcodeTxt
'
'Matchcode$ = Match1.MatchcodeErg
'
'Set Match1 = Nothing
'
'MatchTyp% = -1
'
'Call clsError.DefErrPop
'End Function


Sub ForwardDDE(cmdStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ForwardDDE")
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

frmMatchcode.GetDDE (cmdStr$)

Call clsError.DefErrPop
End Sub

Sub AnzeigeFenster(Typ$, key As Variant, txt$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AnzeigeFenster")
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
Dim lief%
Dim pzn$

AnzeigeFensterTyp$ = Typ$
Load frmBstatus
If (Typ$ = TEXT_BESTELLSTATUS) Then
    pzn$ = key
    Call BstatusBefuellen(pzn$)
ElseIf (Typ$ = TEXT_NACHBEARBEITUNG) Then
    lief% = key
    Call NachbearbeitungBefuellen(lief%)
ElseIf (Typ$ = TEXT_ABSAGEN) Then
    lief% = key
    Call AbsagenBefuellen(lief%)
End If

With frmBstatus
    .Caption = Typ$ + " - " + txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmBstatus.Show 1

Call clsError.DefErrPop
End Sub

Sub BestellStatus(pzn$, txt$)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BestellStatus")
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
Load frmBstatus
Call BstatusBefuellen(pzn$)
'
With frmBstatus
    .Caption = "Bestellstatus - " + txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmBstatus.Show 1

Call clsError.DefErrPop
End Sub

Sub BstatusBefuellen(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BstatusBefuellen")
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
Dim i%, Erst%, max%, lief%, iZeit%, alm1%, arm1%, m%, IstWawiMdb%, dAnzFields%, IstMerkzettelMdb%
Dim FabsRecno&
Dim h$, h2$, nbtext$, SQLStr$

With frmBstatus.flxbstatus
    .row = 0
    
    Erst% = True
    If (ArtikelDBok) Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn$
        SQLStr = SQLStr + " AND LagerKz<>0"
        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
        If (FabsErrf = 0) Then
            SQLStr$ = "SELECT * FROM LIEFERUNGEN WHERE PZN = " + pzn
            SQLStr = SQLStr + " ORDER BY BelegDatum DESC"
            FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
            If (FabsErrf% = 0) Then
                lief% = ArtikelRec2!LiefNr
                h$ = Format(lief%, "###    ")
                If (LieferantenDBok) Then
                    h2 = ""
                    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                    If (LieferantenRec.RecordCount <> 0) Then
                        h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                    End If
                    LieferantenRec.Close
                    h$ = h$ + h2$
                ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                    Lif1.GetRecord (lief% + 1)
                    h2$ = Lif1.kurz
                    Call OemToChar(h2$, h2$)
                    h$ = h$ + h2$
                End If
                If (Erst% = True) Then
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = True
                    .TextMatrix(.row, 0) = "Letzte Lieferung"
                    Erst% = False
                End If
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = False
                .TextMatrix(.row, 1) = Format(ArtikelRec2!BelegDatum, "DD.MM.YY")
                .TextMatrix(.row, 3) = h$
                .TextMatrix(.row, 4) = CStr(ArtikelRec2!lm)
            End If
        End If
    Else
        FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Ass1.GetRecord (FabsRecno& + 1)
            lief% = Ass1.llief(4)
            h$ = Format(lief%, "###    ")
            If (LieferantenDBok) Then
                h2 = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                End If
                LieferantenRec.Close
                h$ = h$ + h2$
            ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                Lif1.GetRecord (lief% + 1)
                h2$ = Lif1.kurz
                Call OemToChar(h2$, h2$)
                h$ = h$ + h2$
            End If
            If (Erst% = True) Then
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = True
                .TextMatrix(.row, 0) = "Letzte Lieferung"
                Erst% = False
            End If
            .Rows = .Rows + 1
            .row = .Rows - 1
            .col = 0
            .CellFontBold = False
            h2$ = clsOpTool.CVDatum2$(Ass1.lld)
            .TextMatrix(.row, 1) = Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2)
            .TextMatrix(.row, 3) = h$
            .TextMatrix(.row, 4) = Ass1.llm
        End If
    End If
    
    If (WawiDBok) Then
        SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn=" + clsOpTool.SqlPzn(pzn$)
        FabsErrf = WawiDB1.OpenRecordset(WawiAdoRec, SQLStr)
        If Not (WawiAdoRec.EOF) Then
            Erst% = True
            WawiAdoRec.MoveFirst
            Do
                If (WawiAdoRec.EOF) Then
                    Exit Do
                End If
                
                If (WawiAdoRec!Status = 1) Then
                    lief% = WawiAdoRec!lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Bestellung"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = Abs(WawiAdoRec!bm) + Abs(WawiAdoRec!nm)
                End If
                
                WawiAdoRec.MoveNext
            Loop
        
            Erst% = True
            WawiAdoRec.MoveFirst
            Do
                If (WawiAdoRec.EOF) Then
                    Exit Do
                End If
                
                If (WawiAdoRec!Status = 2) And (WawiAdoRec!IstAltLast = 0) Then
                    lief% = WawiAdoRec!lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Warenübernahme"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 1) = Left$(WawiAdoRec!WuBestDatum, 2) + "." + Mid$(WawiAdoRec!WuBestDatum, 3, 2) + "." + Mid$(WawiAdoRec!WuBestDatum, 5, 2)
'                    iZeit% = clsOpTool.CVI(WawiAdoRec!WuBestZeit)
                    iZeit% = Val(WawiAdoRec!WuBestZeit)
                    .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = Abs(WawiAdoRec!bm) + Abs(WawiAdoRec!nm)
                End If
                
                WawiAdoRec.MoveNext
            Loop
        
            Erst% = True
            WawiAdoRec.MoveFirst
            Do
                If (WawiAdoRec.EOF) Then
                    Exit Do
                End If
                
                If (WawiAdoRec!Status = 2) And (WawiAdoRec!IstAltLast) Then
                    alm1% = WawiAdoRec!WuNeuLm: arm1% = WawiAdoRec!WuNeuRm: nbtext$ = ""
                    If (alm1% > 0) Then
                        nbtext$ = "Nachlieferung"
                    ElseIf (alm1% < 0) Then
                        nbtext$ = "Retourware"
                    End If
                    
                    lief% = WawiAdoRec!lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Altlasten"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 0) = nbtext$
                    .TextMatrix(.row, 1) = Left$(WawiAdoRec!WuBestDatum, 2) + "." + Mid$(WawiAdoRec!WuBestDatum, 3, 2) + "." + Mid$(WawiAdoRec!WuBestDatum, 5, 2)
'                    iZeit% = clsOpTool.CVI(WawiAdoRec!WuBestZeit)
                    iZeit% = Val(WawiAdoRec!WuBestZeit)
                    .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = alm1%
                End If
                
                WawiAdoRec.MoveNext
            Loop
        End If
    Else
        On Error Resume Next
        dAnzFields% = WawiDB.TableDefs("WAWIDAT").Fields.Count '+ 1
        IstWawiMdb% = (Err = 0)
        On Error GoTo DefErr
        
        If (IstWawiMdb%) Then
    '        SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
            SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn=" + clsOpTool.SqlPzn(pzn$)
            Set WawiRec = WawiDB.OpenRecordset(SQLStr$)
            If (WawiRec.RecordCount > 0) Then
                Erst% = True
                WawiRec.MoveFirst
                Do
                    If (WawiRec.EOF) Then
                        Exit Do
                    End If
                    
                    If (WawiRec!Status = 1) Then
                        lief% = WawiRec!lief
                        h$ = Format(lief%, "###    ")
                        If (LieferantenDBok) Then
                            h2 = ""
                            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                            If (LieferantenRec.RecordCount <> 0) Then
                                h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                            End If
                            LieferantenRec.Close
                            h$ = h$ + h2$
                        ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                            Lif1.GetRecord (lief% + 1)
                            h2$ = Lif1.kurz
                            Call OemToChar(h2$, h2$)
                            h$ = h$ + h2$
                        End If
                        If (Erst% = True) Then
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 0
                            .CellFontBold = True
                            .TextMatrix(.row, 0) = "Bestellung"
                            Erst% = False
                        End If
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = False
                        .TextMatrix(.row, 3) = h$
                        .TextMatrix(.row, 4) = Abs(WawiRec!bm) + Abs(WawiRec!nm)
                    End If
                    
                    WawiRec.MoveNext
                Loop
            
                Erst% = True
                WawiRec.MoveFirst
                Do
                    If (WawiRec.EOF) Then
                        Exit Do
                    End If
                    
                    If (WawiRec!Status = 2) And (WawiRec!IstAltLast = 0) Then
                        lief% = WawiRec!lief
                        h$ = Format(lief%, "###    ")
                        If (LieferantenDBok) Then
                            h2 = ""
                            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                            If (LieferantenRec.RecordCount <> 0) Then
                                h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                            End If
                            LieferantenRec.Close
                            h$ = h$ + h2$
                        ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                            Lif1.GetRecord (lief% + 1)
                            h2$ = Lif1.kurz
                            Call OemToChar(h2$, h2$)
                            h$ = h$ + h2$
                        End If
                        If (Erst% = True) Then
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 0
                            .CellFontBold = True
                            .TextMatrix(.row, 0) = "Warenübernahme"
                            Erst% = False
                        End If
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = False
                        .TextMatrix(.row, 1) = Left$(WawiRec!WuBestDatum, 2) + "." + Mid$(WawiRec!WuBestDatum, 3, 2) + "." + Mid$(WawiRec!WuBestDatum, 5, 2)
                        iZeit% = clsOpTool.CVI(WawiRec!WuBestZeit)
                        .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                        .TextMatrix(.row, 3) = h$
                        .TextMatrix(.row, 4) = Abs(WawiRec!bm) + Abs(WawiRec!nm)
                    End If
                    
                    WawiRec.MoveNext
                Loop
            
            
                Erst% = True
                WawiRec.MoveFirst
                Do
                    If (WawiRec.EOF) Then
                        Exit Do
                    End If
                    
                    If (WawiRec!Status = 2) And (WawiRec!IstAltLast) Then
                        alm1% = WawiRec!WuNeuLm: arm1% = WawiRec!WuNeuRm: nbtext$ = ""
                        If (alm1% > 0) Then
                            nbtext$ = "Nachlieferung"
                        ElseIf (alm1% < 0) Then
                            nbtext$ = "Retourware"
                        End If
                        
                        lief% = WawiRec!lief
                        h$ = Format(lief%, "###    ")
                        If (LieferantenDBok) Then
                            h2 = ""
                            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                            If (LieferantenRec.RecordCount <> 0) Then
                                h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                            End If
                            LieferantenRec.Close
                            h$ = h$ + h2$
                        ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                            Lif1.GetRecord (lief% + 1)
                            h2$ = Lif1.kurz
                            Call OemToChar(h2$, h2$)
                            h$ = h$ + h2$
                        End If
                        If (Erst% = True) Then
                            .Rows = .Rows + 1
                            .row = .Rows - 1
                            .col = 0
                            .CellFontBold = True
                            .TextMatrix(.row, 0) = "Altlasten"
                            Erst% = False
                        End If
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = False
                        .TextMatrix(.row, 0) = nbtext$
                        .TextMatrix(.row, 1) = Left$(WawiRec!WuBestDatum, 2) + "." + Mid$(WawiRec!WuBestDatum, 3, 2) + "." + Mid$(WawiRec!WuBestDatum, 5, 2)
                        iZeit% = clsOpTool.CVI(WawiRec!WuBestZeit)
                        .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                        .TextMatrix(.row, 3) = h$
                        .TextMatrix(.row, 4) = alm1%
                    End If
                    
                    WawiRec.MoveNext
                Loop
            End If
        
        Else
        
            Erst% = True
            Ww1.GetRecord (1)
            max% = Ww1.erstmax
            For i% = 1 To max%
                Ww1.GetRecord (i% + 1)
                If (Ww1.Status = 1) And (pzn$ = Ww1.pzn) Then
                    lief% = Ww1.lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Bestellung"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = Abs(Ww1.bm) + Abs(Ww1.nm)
                End If
            Next i%
            
            
            Erst% = True
            Ww1.GetRecord (1)
            max% = Ww1.erstmax
            For i% = 1 To max%
                Ww1.GetRecord (i% + 1)
                If (Ww1.Status = 2) And (Ww1.IstAltLast = 0) And (pzn$ = Ww1.pzn) Then
                    lief% = Ww1.lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Warenübernahme"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 1) = Left$(Ww1.WuBestDatum, 2) + "." + Mid$(Ww1.WuBestDatum, 3, 2) + "." + Mid$(Ww1.WuBestDatum, 5, 2)
                    iZeit% = clsOpTool.CVI(Ww1.WuBestZeit)
                    .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = Abs(Ww1.bm) + Abs(Ww1.nm)
                End If
            Next i%
            
            Erst% = True
            Ww1.GetRecord (1)
            max% = Ww1.erstmax
            For i% = 1 To max%
                Ww1.GetRecord (i% + 1)
                If (Ww1.Status = 2) And (Ww1.IstAltLast) And (pzn$ = Ww1.pzn) Then
                    alm1% = Ww1.WuNeuLm: arm1% = Ww1.WuNeuRm: nbtext$ = ""
                    If (alm1% > 0) Then
                        nbtext$ = "Nachlieferung"
                    ElseIf (alm1% < 0) Then
                        nbtext$ = "Retourware"
                    End If
                    
                    lief% = Ww1.lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Altlasten"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 0) = nbtext$
                    .TextMatrix(.row, 1) = Left$(Ww1.WuBestDatum, 2) + "." + Mid$(Ww1.WuBestDatum, 3, 2) + "." + Mid$(Ww1.WuBestDatum, 5, 2)
                    iZeit% = clsOpTool.CVI(Ww1.WuBestZeit)
                    .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = alm1%
                End If
            Next i%
        
        End If
    End If
    
    
    Erst% = True
    
    If (ArtikelDBok) Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn$
        SQLStr = SQLStr + " AND LagerKz=0"
        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
        If (FabsErrf = 0) Then
            SQLStr$ = "SELECT * FROM LIEFERUNGEN WHERE PZN = " + pzn
            SQLStr = SQLStr + " ORDER BY BelegDatum DESC"
            FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
            If (FabsErrf% = 0) Then
                Do
                    If (ArtikelRec2.EOF) Then
                        Exit Do
                    End If
                    
                    lief% = ArtikelRec2!LiefNr
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Merkzettel"
                        Erst% = False
                    End If
                    
                    If (ArtikelRec2!lm = 0) Then
                        nbtext$ = "angefragt"
                    Else
                        nbtext$ = "besorgt"
                    End If
                    
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    
                    .TextMatrix(.row, 0) = nbtext$
                    .TextMatrix(.row, 1) = Format(ArtikelRec2!BelegDatum, "DD.MM.YY")
                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = CStr(ArtikelRec2!lm)
                    
                    ArtikelRec2.MoveNext
                Loop
            End If
        End If
    Else
        On Error Resume Next
        dAnzFields% = MerkzettelDB.TableDefs("Merkzettel").Fields.Count
        IstMerkzettelMdb% = (Err = 0)
        On Error GoTo DefErr
        
        Erst% = True
        If (IstMerkzettelMdb%) Then
            SQLStr$ = "SELECT * FROM Merkzettel WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
            Set MerkzettelRec = MerkzettelDB.OpenRecordset(SQLStr$)
            If (MerkzettelRec.RecordCount > 0) Then
                MerkzettelRec.MoveFirst
                Do
                    If (MerkzettelRec.EOF) Then
                        Exit Do
                    End If
                    
                    m% = MerkzettelRec!lm
                    lief% = MerkzettelRec!lief
                    h$ = Format(lief%, "###    ")
                    If (LieferantenDBok) Then
                        h2 = ""
                        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                        If (LieferantenRec.RecordCount <> 0) Then
                            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                        End If
                        LieferantenRec.Close
                        h$ = h$ + h2$
                    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                        Lif1.GetRecord (lief% + 1)
                        h2$ = Lif1.kurz
                        Call OemToChar(h2$, h2$)
                        h$ = h$ + h2$
                    End If
                    If (m% = 0) Then
                        nbtext$ = "angefragt"
                    Else
                        nbtext$ = "besorgt"
                    End If
                    If (Erst% = True) Then
                        .Rows = .Rows + 1
                        .row = .Rows - 1
                        .col = 0
                        .CellFontBold = True
                        .TextMatrix(.row, 0) = "Merkzettel"
                        Erst% = False
                    End If
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = False
                    .TextMatrix(.row, 0) = nbtext$
                    .TextMatrix(.row, 1) = Format(MerkzettelRec!LieferDatum, "DD.MM.YY")
                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 3) = h$
                    .TextMatrix(.row, 4) = m%
        
                    MerkzettelRec.MoveNext
                Loop
            End If
        Else
            FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
            Do
                If (FabsErrf%) Then Exit Do
        
                Besorgt1.GetRecord (FabsRecno& + 1)
        
                If (pzn$ <> Besorgt1.pzn) Then Exit Do
        
                m% = Besorgt1.lm
                lief% = Besorgt1.lief
                h$ = Format(lief%, "###    ")
                If (LieferantenDBok) Then
                    h2 = ""
                    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                    If (LieferantenRec.RecordCount <> 0) Then
                        h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                    End If
                    LieferantenRec.Close
                    h$ = h$ + h2$
                ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                    Lif1.GetRecord (lief% + 1)
                    h2$ = Lif1.kurz
                    Call OemToChar(h2$, h2$)
                    h$ = h$ + h2$
                End If
                If (m% = 0) Then
                    nbtext$ = "angefragt"
                Else
                    nbtext$ = "besorgt"
                End If
                If (Erst% = True) Then
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = True
                    .TextMatrix(.row, 0) = "Merkzettel"
                    Erst% = False
                End If
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = False
                .TextMatrix(.row, 0) = nbtext$
                h2$ = clsOpTool.CVDatum2$(Besorgt1.dt)
        
                .TextMatrix(.row, 1) = Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2)
                .TextMatrix(.row, 2) = ""
                .TextMatrix(.row, 3) = h$
                .TextMatrix(.row, 4) = m%
        
                FabsErrf% = Besorgt1.IndexNext(0, FabsRecno&, pzn$, FabsRecno&)
            Loop
        End If
    End If
    
    If (.Rows = 1) Then
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 0) = "keine Daten gefunden !"
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Sub NachbearbeitungBefuellen(lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("NachbearbeitungBefuellen")
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
Dim i%, Erst%, max%, iZeit%, alm1%, arm1%, m%
Dim FabsRecno&
Dim h$, h2$, nbtext$

With frmBstatus.flxbstatus
    .row = 0
    
    Erst% = True
    Nb1.OpenDatei
    Nb1.GetRecord (1)
    max% = Nb1.erstmax
   
    For i% = 1 To max%
        Nb1.GetRecord (i% + 1)
        If (lief% = Asc(Nb1.li)) Then
            .Rows = .Rows + 1
            .row = .Rows - 1
            .TextMatrix(.row, 0) = Left$(Nb1.bd, 2) + "." + Mid$(Nb1.bd, 3, 2) + "." + Mid$(Nb1.bd, 5, 2)
            .TextMatrix(.row, 1) = Left$(Nb1.at, 7)
            .TextMatrix(.row, 2) = Mid$(Nb1.at, 9, 28)
            .TextMatrix(.row, 3) = Mid$(Nb1.at, 37, 5)
            .TextMatrix(.row, 4) = Mid$(Nb1.at, 42, 3)
            .TextMatrix(.row, 5) = Nb1.bm
            If (Nb1.nm <> 0) Then
                .TextMatrix(.row, 6) = Nb1.nm
            Else
                .TextMatrix(.row, 6) = ""
            End If
            .TextMatrix(.row, 7) = Nb1.rm
            .TextMatrix(.row, 8) = Nb1.lm
            .TextMatrix(.row, 9) = Nb1.user
            
            .Rows = .Rows + 1
            .row = .Rows - 1
            .TextMatrix(.row, 2) = Nb1.txt
            .TextMatrix(.row, 7) = Nb1.bm - Nb1.rm
            .TextMatrix(.row, 8) = Nb1.bm + Nb1.nm - Nb1.lm
        End If
    Next i%
    
    Nb1.CloseDatei
    
    If (.Rows = 1) Then
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 2) = "keine Daten gefunden !"
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Sub AbsagenBefuellen(lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AbsagenBefuellen")
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
Dim i%, max%, iZeit%, alm1%, arm1%, m%
Dim FabsRecno&
Dim h$, h2$, nbtext$
Dim AbsagenRec As New ADODB.Recordset

With frmBstatus.flxbstatus
    .row = 0
    
    If (AbsagenDB1.DBvorhanden) Then
        AbsagenDB1.OpenDB
        SQLStr$ = "SELECT * FROM Absagen WHERE Lief=" + CStr(lief)
        SQLStr = SQLStr + " ORDER BY Datum"
        AbsagenRec.Open SQLStr, AbsagenDB1.ActiveConn
        
        Do
            If (AbsagenRec.EOF) Then
                Exit Do
            End If
            
            .Rows = .Rows + 1
            .row = .Rows - 1
            
            h = clsOpTool.PznString(AbsagenRec!pzn)
            .TextMatrix(.row, 0) = Format(AbsagenRec!datum, "DD.MM.YY")
            .TextMatrix(.row, 1) = h

            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + h
            If (TaxeAdoDBok) Then
                On Error Resume Next
                TaxeAdoRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
                If (TaxeAdoRec.EOF = False) Then
                    .TextMatrix(.row, 2) = Trim(TaxeAdoRec!Name)
                    .TextMatrix(.row, 3) = Trim(TaxeAdoRec!menge)
                    .TextMatrix(.row, 4) = Trim(TaxeAdoRec!einheit)
                End If
            Else
                Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                If (TaxeRec.EOF = False) Then
                    .TextMatrix(.row, 2) = Trim(TaxeRec!Name)
                    .TextMatrix(.row, 3) = Trim(TaxeRec!menge)
                    .TextMatrix(.row, 4) = Trim(TaxeRec!einheit)
                End If
            End If
            
            .TextMatrix(.row, 5) = Str$(AbsagenRec!menge)
            
            AbsagenRec.MoveNext
        Loop
        AbsagenRec.Close
        
        AbsagenDB1.CloseDB
    Else
        Absagen1.OpenDatei
        Absagen1.GetRecord (1)
        max% = Absagen1.erstmax
       
        For i% = 1 To max%
            Absagen1.GetRecord (i% + 1)
            If (lief% = Absagen1.lief) Then
                .Rows = .Rows + 1
                .row = .Rows - 1
                
                h2$ = clsOpTool.CVDatum2$(Absagen1.datum)
                .TextMatrix(.row, 0) = Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2)
                .TextMatrix(.row, 1) = Absagen1.pzn
    
                h2$ = Absagen1.text
                .TextMatrix(.row, 2) = Left$(h2$, 28)
                .TextMatrix(.row, 3) = Mid$(h2$, 29, 5)
                .TextMatrix(.row, 4) = Mid$(h2$, 34, 3)
                
                .TextMatrix(.row, 5) = Absagen1.menge
            End If
        Next i%
        
        Absagen1.CloseDatei
    End If
    
    If (.Rows = 1) Then
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 2) = "keine Daten gefunden !"
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Function TesteBstatus$(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TesteBstatus$")
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
Dim i%, max%, gef%, IstWawiMdb%, dAnzFields%, IstMerkzettelMdb%
Dim FabsRecno&
Dim ret$, SQLStr$

ret$ = Space$(3)
    
'FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
'If (FabsErrf% = 0) Then
'    Call clsError.DefErrPop: Exit Function
'End If
    
'gef% = False
'Bek1.GetRecord (1)
'Max% = Bek1.erstmax
'For i% = 1 To Max%
'    Bek1.GetRecord (i% + 1)
'    If (pzn$ = Bek1.pzn) Then
'        ret$ = ret$ + "B"
'        gef% = True
'        Exit For
'    End If
'Next i%
'If (gef% = False) Then ret$ = ret$ + " "
    
gef% = False
If (WawiDBok) Then
    SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn=" + clsOpTool.SqlPzn(pzn$)
    FabsErrf = WawiDB1.OpenRecordset(WawiAdoRec, SQLStr)
    If Not (WawiAdoRec.EOF) Then
        Do
            If (WawiAdoRec.EOF) Then
                Exit Do
            End If
            
            If (WawiAdoRec!Status = 1) Then
                Mid$(ret$, 1, 1) = "B"
            ElseIf (WawiAdoRec!Status = 2) And (WawiAdoRec!IstAltLast = 0) Then
                Mid$(ret$, 2, 1) = "W"
            ElseIf (WawiAdoRec!Status = 2) And (WawiAdoRec!IstAltLast) Then
                Mid$(ret$, 3, 1) = "N"
            End If
            
            WawiAdoRec.MoveNext
        Loop
    End If
Else
    On Error Resume Next
    dAnzFields% = WawiDB.TableDefs("WAWIDAT").Fields.Count '+ 1
    IstWawiMdb% = (Err = 0)
    On Error GoTo DefErr
    
    If (IstWawiMdb%) Then
    '    SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
        SQLStr$ = "SELECT * FROM WawiDat WHERE Pzn=" + clsOpTool.SqlPzn(pzn$)
        Set WawiRec = WawiDB.OpenRecordset(SQLStr$)
        Do
            If (WawiRec.EOF) Then
                Exit Do
            End If
            
            If (WawiRec!Status = 1) Then
                Mid$(ret$, 1, 1) = "B"
            ElseIf (WawiRec!Status = 2) And (WawiRec!IstAltLast = 0) Then
                Mid$(ret$, 2, 1) = "W"
            ElseIf (WawiRec!Status = 2) And (WawiRec!IstAltLast) Then
                Mid$(ret$, 3, 1) = "N"
            End If
            
            WawiRec.MoveNext
        Loop
    Else
        Ww1.GetRecord (1)
        max% = Ww1.erstmax
        For i% = 1 To max%
            Ww1.GetRecord (i% + 1)
            If (pzn$ = Ww1.pzn) Then
                If (Ww1.Status = 1) Then
                    Mid$(ret$, 1, 1) = "B"
                ElseIf (Ww1.Status = 2) And (Ww1.IstAltLast = 0) Then
                    Mid$(ret$, 2, 1) = "W"
                ElseIf (Ww1.Status = 2) And (Ww1.IstAltLast) Then
                    Mid$(ret$, 3, 1) = "N"
                End If
            End If
        Next i%
    End If
End If
    
'gef% = False
'Call Wu1.GetRecord(1)
'max% = Wu1.erstmax
'For i% = 1 To max%
'    Call Wu1.GetRecord(i% + 1)
'    If (pzn$ = Left$(Wu1.at, 7)) Then
'        ret$ = ret$ + "W"
'        gef% = True
'        Exit For
'    End If
'Next i%
'If (gef% = False) Then ret$ = ret$ + " "
'
'gef% = False
'Nb1.OpenDatei
'Nb1.GetRecord (1)
'max% = Nb1.erstmax
'For i% = 1 To max%
'    Nb1.GetRecord (i% + 1)
'    If (pzn$ = Left$(Nb1.at, 7)) Then
'        ret$ = ret$ + "N"
'        gef% = True
'        Exit For
'    End If
'Next i%
'Nb1.CloseDatei
'If (gef% = False) Then ret$ = ret$ + " "

gef% = False

If (ArtikelDBok) Then
    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn$
    SQLStr = SQLStr + " AND Lagerkz=0"
    FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
    If (FabsErrf = 0) Then
        ret$ = ret$ + "M"
        gef% = True
    End If
Else
    On Error Resume Next
    dAnzFields% = MerkzettelDB.TableDefs("Merkzettel").Fields.Count
    IstMerkzettelMdb% = (Err = 0)
    On Error GoTo DefErr
    
    If (IstMerkzettelMdb%) Then
        SQLStr$ = "SELECT * FROM Merkzettel WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
        Set MerkzettelRec = MerkzettelDB.OpenRecordset(SQLStr$)
        If (MerkzettelRec.RecordCount > 0) Then
            ret$ = ret$ + "M"
            gef% = True
        End If
    Else
        FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ret$ = ret$ + "M"
            gef% = True
        End If
    End If
End If
If (gef% = False) Then ret$ = ret$ + " "
    
TesteBstatus$ = ret$

Call clsError.DefErrPop
End Function

Function ZusatzFenster%(Typ$, key As Variant, txt$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZusatzFenster%")
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
Dim pzn$

ZusatzFensterTyp$ = Typ$
ZusatzChanged% = False
Load frmZusatz
If (Typ$ = ZUSATZ_ARTIKEL) Then
    pzn$ = key
    ZusatzPzn$ = pzn$
    If (ArtikelDBok) Then
        Call ArtikelDB1.ZusatzBefuellen(pzn)
    Else
        Call Zus1.ZusatzBefuellen(pzn$)
    End If
ElseIf (Typ$ = ZUSATZ_LIEFERANTEN) Then
    pzn$ = Format(key, "00000")
    ZusatzPzn$ = pzn$
    Call lZus1.ZusatzBefuellen(pzn$)
'ElseIf (Typ$ = ZUSATZ_KUNDEN) Then
'    pzn$ = Format(key, "00000")
'    ZusatzPzn$ = pzn$
'    Call kZus1.ZusatzBefuellen(pzn$)
End If

With frmZusatz
'    Err.Raise 7
    .Caption = txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmZusatz.Show 1

ZusatzFenster% = ZusatzChanged%

Call clsError.DefErrPop
End Function

Function MyInputBox$(prompt$, Optional title$ = "", Optional default$ = "", Optional iEditModus% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MyInputBox$")
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

If (iNewLine) And (nlmsgAktiv = 0) Then
    ret = nlInputBox(prompt$, title$, default$, iEditModus)
Else
    ret = InputBox(prompt$, title$, default$)
End If

MyInputBox$ = ret$

Call clsError.DefErrPop
End Function

'Function nlMsgBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("nlMsgBox%")
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
'Dim ret%
'
'nlmsgDefault(0) = True
'nlmsgCancel(0) = True
'nlmsgVisible(0) = True
'nlmsgCaption(0) = "OK"
'
'nlmsgDefault(1) = 0
'nlmsgCancel(1) = 0
'nlmsgVisible(1) = 0
'
'If (buttons And vbYesNo) Then
'    nlmsgDefault(0) = True
'    nlmsgCancel(0) = 0
'    nlmsgVisible(0) = True
'    nlmsgCaption(0) = "Ja"
'
'    nlmsgDefault(1) = 0
'    nlmsgCancel(1) = True
'    nlmsgVisible(1) = True
'    nlmsgCaption(1) = "Nein"
'
'    If (buttons And vbDefaultButton2) Then
'        nlmsgDefault(0) = 0
'        nlmsgDefault(1) = True
'    End If
'End If
'
'nlmsgPicto = -1
'If (buttons And vbYesNo) Then
'    nlmsgPicto = 0
'ElseIf (buttons And vbInformation) Then
'    nlmsgPicto = 1
'ElseIf (buttons And vbExclamation) Then
'    nlmsgPicto = 2
'ElseIf (buttons And vbCritical) Then
'    nlmsgPicto = 2
'End If
'
'nlmsgTitle = Trim(title)
'If (nlmsgTitle = "") Then
'    nlmsgTitle = App.EXEName
'End If
'
'nlmsgPrompt = prompt
'
'frmnlMsg.Show 1
'
'nlMsgBox = nlmsgRet
'
''ret% = nlMsgBox(prompt$, buttons%, title$)
'
'Call clsError.DefErrPop
'End Function

Function MessageBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MessageBox%")
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
Dim ret%

If (iNewLine) And (nlmsgAktiv = 0) Then
    ret% = nlMsgBox(prompt$, buttons%, title$)
Else
    If (title$ <> "") Then
        ret% = MsgBox(prompt$, buttons%, title$)
    Else
        ret% = MsgBox(prompt$, buttons%)
    End If
End If

MessageBox% = ret%

Call clsError.DefErrPop
End Function

Function nlMsgBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("nlMsgBox%")
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
Dim ret%

nlmsgDefault(0) = True
nlmsgCancel(0) = True
nlmsgVisible(0) = True
nlmsgCaption(0) = "&OK"

nlmsgDefault(1) = 0
nlmsgCancel(1) = 0
nlmsgVisible(1) = 0

nlmsgCaption(2) = ""
nlmsgVisible(2) = 0

If (buttons And vbYesNo) Or (buttons And vbOKCancel) Then
    nlmsgDefault(0) = True
    nlmsgCancel(0) = 0
    nlmsgVisible(0) = True
    nlmsgCaption(0) = "&Ja"
    
    nlmsgDefault(1) = 0
    nlmsgCancel(1) = True
    nlmsgVisible(1) = True
    nlmsgCaption(1) = "&Nein"

    If (buttons And vbDefaultButton2) Then
        nlmsgDefault(0) = 0
        nlmsgDefault(1) = True
    End If
End If

nlmsgPicto = -1
If (buttons And vbOKCancel) Then
    nlmsgPicto = 0
    nlmsgCaption(0) = "&Ok"
    nlmsgCaption(1) = "&Abbruch"
ElseIf (buttons And vbYesNo) Then
    nlmsgPicto = 0
ElseIf (buttons And vbInformation) Then
    nlmsgPicto = 1
ElseIf (buttons And vbExclamation) Then
    nlmsgPicto = 2
ElseIf (buttons And vbCritical) Then
    nlmsgPicto = 2
End If

nlmsgTitle = Trim(title)
If (nlmsgTitle = "") Then
    nlmsgTitle = App.EXEName
End If

nlmsgPrompt = prompt

frmnlMsg.Show 1

nlMsgBox = nlmsgRet

'ret% = nlMsgBox(prompt$, buttons%, title$)

Call clsError.DefErrPop
End Function

Function nlInputBox$(prompt$, Optional title$ = "", Optional default = "", Optional iEditModus% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("nlInputBox$")
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
Dim ret%

nlmsgDefault(0) = True
nlmsgCancel(0) = 0
nlmsgVisible(0) = True
nlmsgCaption(0) = "&OK"

nlmsgDefault(1) = 0
nlmsgCancel(1) = True
nlmsgVisible(1) = True
nlmsgCaption(1) = "&ESC"

nlmsgCaption(2) = default
nlmsgVisible(2) = True
nlmsgEditModus = iEditModus

nlmsgPicto = -1

nlmsgTitle = Trim(title)
If (nlmsgTitle = "") Then
    nlmsgTitle = App.EXEName
End If

nlmsgPrompt = prompt

frmnlMsg.Show 1

nlInputBox = nlmsgRet2

'ret% = nlMsgBox(prompt$, buttons%, title$)

Call clsError.DefErrPop
End Function


Sub BesorgerInfo(Nummer%, pzn$, txt$, Optional bModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BesorgerInfo")
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
Dim erg%, Belegt%, kStatus%, iOk%
Dim ErrNumber&
Dim h$, SQLStr$, sSqlDatabase$

If (Nummer% < 0) Then
    Call clsError.DefErrPop: Exit Sub
End If

'BesorgerBenutzer% = 0
'If (bModus% >= 1000) Then
'    BesorgerBenutzer% = bModus% - 1000
'    bModus% = 0
'End If
BesorgerBenutzer% = Para1.AktBenutzer

AbholerMdb% = 0

sSqlDatabase = "Abholer" + sqlop1.SqlGetSuffix
AbholerSQL = sqlop1.SqlCheckDatabase(sSqlDatabase)
If (AbholerSQL) Then
    AbholerMdb% = True
ElseIf (Dir(ABHOLER_DB) <> "") Then
    On Error Resume Next
    Err.Clear
    Set AbholerDB = OpenDatabase(ABHOLER_DB, False, True)
    If (Err.Number = 0) Then
        AbholerMdb% = True
        AbholerDB.Close
    End If
    Err.Clear
    On Error GoTo DefErr
End If


If (AbholerSQL) Then
    h = sqlop1.SqlConnectionString(AbholerSQL)
    AbholerConn.ConnectionString = Left(h, Len(h) - 1) + sSqlDatabase + "; Data Source=" + sqlop1.SqlServer(AbholerSQL)

    AbholerConn.CursorLocation = adUseClient
    AbholerConn.CommandTimeout = 300
    On Error Resume Next
    Err.Clear
    AbholerConn.Open
    ErrNumber = Err.Number
    On Error GoTo DefErr
    erg = (ErrNumber = 0)
    If (erg = 0) Then
    End If

    SQLStr$ = "SELECT * FROM AbholerNummern WHERE AbholerNummer=" + Str$(Nummer%)
    On Error Resume Next
    AbholerNummerAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    AbholerNummerAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
    iOk = Not (AbholerNummerAdoRec.EOF)
    If (iOk%) Then
        kStatus% = clsOpTool.CheckNullLong(AbholerNummerAdoRec!Status)
        iOk% = (kStatus% >= 1) And (kStatus% <= 3)
    End If
    If (iOk%) Then
        BesorgerNummer% = Nummer%
        BesorgerPzn$ = pzn$
        BesorgerModus% = bModus%
        Load frmAbholer
        With frmAbholer
            h$ = "Abholnummer " + Str$(Nummer%)
            If (AbholerNummerAdoRec!Status = 3) Then
                h$ = h$ + " - " + "FERTIG"
            ElseIf (AbholerNummerAdoRec!Status = 4) Then
                h$ = h$ + " - " + "ABGEHOLT"
            End If
            .Caption = h$
    '        .Caption = "Abholnummer " + Str$(Nummer%) + " - " + txt$
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
            .Show 1
            If (BesorgerModus% = 2) Then
                bModus% = 2
            End If
        End With
    End If

    AbholerConn.Close
ElseIf (AbholerMdb%) Then
    Set AbholerDB = OpenDatabase(ABHOLER_DB, False, False)
        
    Set AbholerParaRec = AbholerDB.OpenRecordset("Parameter", dbOpenTable)
    AbholerParaRec.Index = "Unique"
    
    SQLStr$ = "SELECT * FROM AbholerNummern WHERE AbholerNummer=" + Str$(Nummer%)
    Set AbholerNummerRec = AbholerDB.OpenRecordset(SQLStr$)
    iOk% = (AbholerNummerRec.RecordCount > 0)
    If (iOk%) Then
        kStatus% = clsOpTool.CheckNullLong(AbholerNummerRec!Status)
        iOk% = (kStatus% >= 1) And (kStatus% <= 3)
    End If
    If (iOk%) Then
        BesorgerNummer% = Nummer%
        BesorgerPzn$ = pzn$
        BesorgerModus% = bModus%
        Load frmAbholer
        With frmAbholer
            h$ = "Abholnummer " + Str$(Nummer%)
            If (AbholerNummerRec!Status = 3) Then
                h$ = h$ + " - " + "FERTIG"
            ElseIf (AbholerNummerRec!Status = 4) Then
                h$ = h$ + " - " + "ABGEHOLT"
            End If
            .Caption = h$
    '        .Caption = "Abholnummer " + Str$(Nummer%) + " - " + txt$
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
            .Show 1
            If (BesorgerModus% = 2) Then
                bModus% = 2
            End If
        End With
    End If
        
    AbholerDB.Close
Else
    If (Abs(Nummer%) > 999) Then
        Call clsError.DefErrPop: Exit Sub
    End If
    
    Set Kiste1 = New clsKiste
    Kiste1.OpenDatei
    Belegt% = Kiste1.Belegt(Nummer%)
    If (Belegt%) Then
        BesorgerNummer% = Nummer%
        BesorgerPzn$ = pzn$
        BesorgerModus% = bModus%
        Load frmAbholer
        With frmAbholer
            h$ = "Abholnummer " + Str$(Nummer%)
            If (Kiste1.KistenStatus = 3) Then
                h$ = h$ + " - " + "FERTIG"
            ElseIf (Kiste1.KistenStatus = 4) Then
                h$ = h$ + " - " + "ABGEHOLT"
            End If
            .Caption = h$
    '        .Caption = "Abholnummer " + Str$(Nummer%) + " - " + txt$
            .Left = (Screen.Width - .Width) / 2
            .Top = (Screen.Height - .Height) / 2
            .Show 1
            If (BesorgerModus% = 2) Then
                bModus% = 2
            End If
        End With
    End If
    Kiste1.CloseDatei
    Set Kiste1 = Nothing
End If

Call clsError.DefErrPop
End Sub

Sub BesorgerBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BesorgerBefuellen")
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
Dim i%, erg%, ret%, PersCode%, iRow%
Dim KundenNr&
Dim h$, pzn$, tx$, RezTxt$, SQLStr$

If (AbholerSQL) Then
    PersCode% = AbholerNummerAdoRec!AngelegtVon
    KundenNr = AbholerNummerAdoRec!KuNr
    
    With frmAbholer
        tx$ = ""
        If (PersCode% > 0) Then
            tx$ = Para1.Personal(PersCode%)
        End If
        .flxAbholerGlobal.TextMatrix(0, 1) = tx$
        
        tx$ = AbholerNummerAdoRec!AngelegtComputer
        If (AbholerNummerAdoRec!AngelegtLiRe = 1) Then
            tx$ = tx$ + " Rechts"
        Else
            tx$ = tx$ + " Links"
        End If
        .flxAbholerGlobal.TextMatrix(1, 1) = tx$
        
        tx$ = Format(AbholerNummerAdoRec!AngelegtAm, "DD.MM.YYYY")
        .flxAbholerGlobal.TextMatrix(2, 1) = tx$
        
        tx$ = Format(AbholerNummerAdoRec!AngelegtAm, "HH:MM")
        .flxAbholerGlobal.TextMatrix(3, 1) = tx$
        
        tx$ = ""
        If (KundenNr > 0) Then
            tx$ = clsOpTool.HoleKundenName(KundenNr)
        End If
        .flxAbholerGlobal.TextMatrix(4, 1) = tx$
        
        tx$ = ""
        On Error Resume Next
        tx$ = Trim(AbholerNummerAdoRec!LieferScheinNr)
        On Error GoTo DefErr
        .txtLieferschein.Visible = (tx$ <> "")
'        .txtLieferschein.Visible = (Trim(AbholerNummerRec!LieferScheinNr) <> "")
        
        iRow% = 0
        
        SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
        On Error Resume Next
        AbholerDetailAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        AbholerDetailAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
        
        For i% = 0 To 50
            If (AbholerDetailAdoRec.EOF) Then
                Exit For
            End If
            
            pzn$ = clsOpTool.PznString(AbholerDetailAdoRec!pzn)
                
            If (Val(pzn$) = 9999999) Then
            Else
                If (pzn$ = BesorgerPzn$) Then
                    iRow% = i%
                End If
            End If
            h$ = Trim(AbholerDetailAdoRec!text) + "  " + Trim(AbholerDetailAdoRec!menge) + " " + Trim(AbholerDetailAdoRec!einheit)
            .flxAbholerGlobal.TextMatrix(5 + i%, 1) = h$
            If (AbholerDetailAdoRec!Status = 3) Then
                .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "FERTIG"
            ElseIf (AbholerDetailAdoRec!Status = 4) Then
                .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "ABGEHOLT"
            End If
            
            AbholerDetailAdoRec.MoveNext
        Next i%
        .flxAbholerGlobal.Rows = 5 + i%
        
        .flxAbholerGlobal.row = 5 + iRow%
        .flxAbholerGlobal.ColSel = .flxAbholerGlobal.Cols - 1
        Call ZeigeAbholerEinzeln(iRow%)
    End With
ElseIf (AbholerMdb%) Then
    PersCode% = AbholerNummerRec!AngelegtVon
    KundenNr = AbholerNummerRec!KuNr
    
    With frmAbholer
        tx$ = ""
        If (PersCode% > 0) Then
            tx$ = Para1.Personal(PersCode%)
        End If
        .flxAbholerGlobal.TextMatrix(0, 1) = tx$
        
        tx$ = AbholerNummerRec!AngelegtComputer
        If (AbholerNummerRec!AngelegtLiRe = 1) Then
            tx$ = tx$ + " Rechts"
        Else
            tx$ = tx$ + " Links"
        End If
        .flxAbholerGlobal.TextMatrix(1, 1) = tx$
        
        tx$ = Format(AbholerNummerRec!AngelegtAm, "DD.MM.YYYY")
        .flxAbholerGlobal.TextMatrix(2, 1) = tx$
        
        tx$ = Format(AbholerNummerRec!AngelegtAm, "HH:MM")
        .flxAbholerGlobal.TextMatrix(3, 1) = tx$
        
        tx$ = ""
        If (KundenNr > 0) Then
            tx$ = clsOpTool.HoleKundenName(KundenNr)
        End If
        .flxAbholerGlobal.TextMatrix(4, 1) = tx$
        
        tx$ = ""
        On Error Resume Next
        tx$ = Trim(AbholerNummerRec!LieferScheinNr)
        On Error GoTo DefErr
        .txtLieferschein.Visible = (tx$ <> "")
'        .txtLieferschein.Visible = (Trim(AbholerNummerRec!LieferScheinNr) <> "")
        
        iRow% = 0
        
        SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerRec!AbholerNummer)
        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
        Set AbholerDetailRec = AbholerDB.OpenRecordset(SQLStr$)
        
        For i% = 0 To 50
            If (AbholerDetailRec.EOF) Then
                Exit For
            End If
            
            pzn$ = clsOpTool.PznString(AbholerDetailRec!pzn)
                
            If (Val(pzn$) = 9999999) Then
            Else
                If (pzn$ = BesorgerPzn$) Then
                    iRow% = i%
                End If
            End If
            h$ = Trim(AbholerDetailRec!text) + "  " + Trim(AbholerDetailRec!menge) + " " + Trim(AbholerDetailRec!einheit)
            .flxAbholerGlobal.TextMatrix(5 + i%, 1) = h$
            If (AbholerDetailRec!Status = 3) Then
                .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "FERTIG"
            ElseIf (AbholerDetailRec!Status = 4) Then
                .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "ABGEHOLT"
            End If
            
            AbholerDetailRec.MoveNext
        Next i%
        .flxAbholerGlobal.Rows = 5 + i%
        
        .flxAbholerGlobal.row = 5 + iRow%
        .flxAbholerGlobal.ColSel = .flxAbholerGlobal.Cols - 1
        Call ZeigeAbholerEinzeln(iRow%)
    End With
Else
    Call Kiste1.GetKiste(BesorgerNummer%)
        
    Call SucheBesorgerInVk(BesorgerNummer%, PersCode%, KundenNr)
    
    With frmAbholer
        tx$ = ""
        If (PersCode% > 0) Then
            tx$ = Para1.Personal(PersCode%)
        End If
        .flxAbholerGlobal.TextMatrix(0, 1) = tx$
        
        tx$ = Kiste1.VonWo \ 2
        If (Kiste1.VonWo Mod 2) Then
            tx$ = tx$ + " Rechts"
        Else
            tx$ = tx$ + " Links"
        End If
        .flxAbholerGlobal.TextMatrix(1, 1) = tx$
        
        tx$ = clsOpTool.CVDatum(Left$(Kiste1.VonWann, 2))
        .flxAbholerGlobal.TextMatrix(2, 1) = Mid$(tx$, 7, 2) + "." + Mid$(tx$, 5, 2) + "." + Left$(tx$, 4)
        
        .flxAbholerGlobal.TextMatrix(3, 1) = Format(Asc(Mid$(Kiste1.VonWann, 3, 1)), "00") + ":" + Format(Asc(Mid$(Kiste1.VonWann, 4, 1)), "00")
        
        tx$ = ""
        If (KundenNr > 0) Then
            tx$ = clsOpTool.HoleKundenName(KundenNr)
        End If
        .flxAbholerGlobal.TextMatrix(4, 1) = tx$
        
        .txtLieferschein.Visible = Kiste1.HatLieferSchein
        
        iRow% = 0
        For i% = 0 To 9
            Kiste1.GetInhalt (i%)
            If (Kiste1.pzn <> "") Then
                h$ = Kiste1.pzn
                If (h$ = "9999999") Then
                    If (Kiste1.WasTun = "B") Then
                        h$ = Trim(Kiste1.InfoText(0))
                    Else
                        h$ = "Rezeptur"
                    End If
                Else
                    If (h$ = BesorgerPzn$) Then iRow% = i%
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(h$)
                    If (TaxeAdoDBok) Then
                        On Error Resume Next
                        TaxeAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
                        If (TaxeAdoRec.EOF = False) Then
                            h$ = Trim(TaxeAdoRec!Name) + "  " + Trim(TaxeAdoRec!menge) + " " + Trim(TaxeAdoRec!einheit)
                        End If
                    Else
                        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        If (TaxeRec.EOF = False) Then
                            h$ = Trim(TaxeRec!Name) + "  " + Trim(TaxeRec!menge) + " " + Trim(TaxeRec!einheit)
                        End If
                    End If
                End If
                .flxAbholerGlobal.TextMatrix(5 + i%, 1) = h$
                If (Kiste1.Status = 3) Then
                    .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "FERTIG"
                ElseIf (Kiste1.Status = 4) Then
                    .flxAbholerGlobal.TextMatrix(5 + i%, 2) = "ABGEHOLT"
                End If
            Else
                .flxAbholerGlobal.Rows = 5 + i%
                Exit For
            End If
        Next i%
        
        .flxAbholerGlobal.row = 5 + iRow%
        .flxAbholerGlobal.ColSel = .flxAbholerGlobal.Cols - 1
        Call ZeigeAbholerEinzeln(iRow%)
        
    End With
End If

Call clsError.DefErrPop
End Sub

Function BesorgerFertig%(Nummer%, pzn$, bStatus%, Optional bMenge% = 0, Optional bPreis# = 0#, Optional bPzn$ = "", Optional ChargenNr$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BesorgerFertig%")
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
Dim i%, iOk%, erg%, Belegt%, KisteFertig%, KisteStatus%, BesorgerStatus%, s%, NurFertigeBesorger%, kInd%, NeuAbholerInd%, NeuInfoInd%, kStatus%
Dim DetailId&, ErrNumber&, lRecs&, NewId&
Dim h$, Arbeit$, SQLStr$, sSqlDatabase$
Dim Fld As Field
Dim besRec2 As Recordset
Dim besAdoRec2 As New ADODB.Recordset

'nur true wenn MachAuftrag
BesorgerFertig = False

If (Nummer% < 0) Then
    Call clsError.DefErrPop: Exit Function
End If

AbholerMdb% = 0

sSqlDatabase = "Abholer" + sqlop1.SqlGetSuffix
AbholerSQL = sqlop1.SqlCheckDatabase(sSqlDatabase)
If (AbholerSQL) Then
    AbholerMdb% = True
ElseIf (Dir(ABHOLER_DB) <> "") Then
    On Error Resume Next
    Err.Clear
    Set AbholerDB = OpenDatabase(ABHOLER_DB, False, True)
    If (Err.Number = 0) Then
        AbholerMdb% = True
        AbholerDB.Close
    End If
    Err.Clear
    On Error GoTo DefErr
End If


If (AbholerSQL) Then
    h = sqlop1.SqlConnectionString(AbholerSQL)
    AbholerConn.ConnectionString = Left(h, Len(h) - 1) + sSqlDatabase + "; Data Source=" + sqlop1.SqlServer(AbholerSQL)

    AbholerConn.CursorLocation = adUseClient
    AbholerConn.CommandTimeout = 300
    On Error Resume Next
    Err.Clear
    AbholerConn.Open
    ErrNumber = Err.Number
    On Error GoTo DefErr
    erg = (ErrNumber = 0)
    If (erg = 0) Then
    End If

    SQLStr$ = "SELECT * FROM AbholerNummern WHERE AbholerNummer=" + Str$(Nummer%)
    On Error Resume Next
    AbholerNummerAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    AbholerNummerAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
    iOk = Not (AbholerNummerAdoRec.EOF)
    If (iOk%) Then
        kStatus% = clsOpTool.CheckNullLong(AbholerNummerAdoRec!Status)
        iOk% = (kStatus% >= 1) And (kStatus% <= 3)
    End If
    If (iOk%) Then
        SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
        On Error Resume Next
        AbholerDetailAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        AbholerDetailAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
        
        DetailId& = -1
        For i% = 0 To 50
            If (AbholerDetailAdoRec.EOF) Then
                Exit For
            End If
            
            If (AbholerDetailAdoRec!pzn = Val(pzn$)) Then
                If (AbholerDetailAdoRec!Status < 3) Then
                    DetailId& = AbholerDetailAdoRec!Id
                    Exit For
                ElseIf (DetailId& < 0) Then
                    DetailId& = AbholerDetailAdoRec!Id
                End If
            End If
            
            AbholerDetailAdoRec.MoveNext
        Next i%
        
        If (DetailId& >= 0) Then
            BesorgerBenutzer% = Para1.AktBenutzer
            
            SQLStr$ = "SELECT * FROM AbholerDetails WHERE Id=" + Str$(DetailId&)
            On Error Resume Next
            AbholerDetailAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            AbholerDetailAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
            If (AbholerDetailAdoRec!pzn = Val(pzn$)) Then
'                AbholerDetailRec.Edit
                
                If (bStatus% And 1) Then
                    AbholerDetailAdoRec!lm = bMenge%
                    If (Val(AbholerDetailAdoRec!RezeptNr) = 0) Then
'                        AbholerDetailRec!Offen = AbholerDetailRec!Preis * AbholerDetailRec!lm - clsOpTool.xVal(AbholerDetailRec!Aconto)
                        AbholerDetailAdoRec!Offen = clsOpTool.FNX(AbholerDetailAdoRec!Preis * AbholerDetailAdoRec!lm - AbholerDetailAdoRec!Aconto)
                    End If
                End If
                If (bStatus% And 2) Then
                    AbholerDetailAdoRec!Preis = bPreis#
                    If (Val(AbholerDetailAdoRec!RezeptNr) = 0) Then
'                        AbholerDetailRec!Offen = AbholerDetailRec!Preis * AbholerDetailRec!lm - clsOpTool.xVal(AbholerDetailRec!Aconto)
                        AbholerDetailAdoRec!Offen = clsOpTool.FNX(AbholerDetailAdoRec!Preis * AbholerDetailAdoRec!lm - AbholerDetailAdoRec!Aconto)
                    End If
                End If
                If (bStatus% And 4) And (AbholerDetailAdoRec!Status <> 4) Then
                    Dim eRezTaskId As String
                    eRezTaskId = clsOpTool.CheckNullStr(AbholerDetailAdoRec!eRezTaskId)

                    If (AbholerDetailAdoRec!lm < AbholerDetailAdoRec!bm) Then
                        NeuInfoInd% = 1
                        SQLStr$ = "SELECT max(InfoInd) as MaxInfoInd FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailAdoRec!Id)
                        On Error Resume Next
                        besAdoRec2.Close
                        Err.Clear
                        On Error GoTo DefErr
                        besAdoRec2.Open SQLStr, AbholerConn ', adOpenDynamic, adLockOptimistic
                        If Not (besAdoRec2.EOF) Then
                            On Error Resume Next
                            NeuInfoInd% = besAdoRec2!MaxInfoInd + 1
                            On Error GoTo DefErr
                        End If
                        
                        lRecs = 0
                        SQLStr = "INSERT INTO AbholerInfo (Status) VALUES (" + CStr(5) + ")"
                        Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                        If (lRecs = 1) Then
                            On Error Resume Next
                            AbholerInfoAdoRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                            AbholerInfoAdoRec.Open SQLStr$, AbholerConn
                            If Not (AbholerInfoAdoRec.EOF) Then
                                NewId& = clsOpTool.CheckNullLong(AbholerInfoAdoRec!NewId)
                            End If
                        
                            AbholerInfoAdoRec.Close
                            SQLStr$ = "SELECT * FROM AbholerInfo WHERE Id=" + CStr(NewId)
                            AbholerInfoAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                        
                            AbholerInfoAdoRec!AbholerDetailId = AbholerDetailAdoRec!Id ' AbholerDetailId&
                            AbholerInfoAdoRec!InfoInd = NeuInfoInd%
                            AbholerInfoAdoRec!Status = 2
                            AbholerInfoAdoRec!AngelegtAm = Now
                            AbholerInfoAdoRec!AngelegtVon = BesorgerBenutzer%
                            AbholerInfoAdoRec!AngelegtComputer = Val(Para1.user)
                            AbholerInfoAdoRec!AngelegtLiRe = 0
                            AbholerInfoAdoRec!text = "Hinweis: Teil-Lieferung!"
                            
                            AbholerInfoAdoRec.Update
                        End If
                        
                        
                        NeuAbholerInd% = 1
                        SQLStr$ = "SELECT max(AbholerInd) as MaxAbholerInd FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
                        On Error Resume Next
                        besAdoRec2.Close
                        Err.Clear
                        On Error GoTo DefErr
                        besAdoRec2.Open SQLStr, AbholerConn ', adOpenDynamic, adLockOptimistic
                        If Not (besAdoRec2.EOF) Then
                            On Error Resume Next
                            NeuAbholerInd% = besAdoRec2!MaxAbholerInd + 1
                            On Error GoTo DefErr
                        End If
                        
                        SQLStr$ = "SELECT * INTO #tmp FROM AbholerDetails WHERE Id=" + Str$(AbholerDetailAdoRec!Id)
                        SQLStr$ = SQLStr$ + ";"
                        SQLStr = SQLStr + "ALTER TABLE #tmp DROP Column Id" + ";"
                        SQLStr = SQLStr + "INSERT INTO AbholerDetails SELECT * FROM #tmp" + ";"
                        SQLStr = SQLStr + "DROP Table #tmp" + ";"
                        Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
    
                        On Error Resume Next
                        besAdoRec2.Close
                        Err.Clear
                        On Error GoTo DefErr
                        SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                        besAdoRec2.Open SQLStr$, AbholerConn
                        If Not (besAdoRec2.EOF) Then
                            NewId& = clsOpTool.CheckNullLong(besAdoRec2!NewId)
                        End If
                    
                        besAdoRec2.Close
                        SQLStr$ = "SELECT * FROM AbholerDetails WHERE Id=" + CStr(NewId)
                        besAdoRec2.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                        
                        besAdoRec2!AbholerInd = NeuAbholerInd%
                        besAdoRec2!bm = AbholerDetailAdoRec!bm - AbholerDetailAdoRec!lm
                        besAdoRec2!lm = besAdoRec2!bm
                        besAdoRec2!Status = 1
                        besAdoRec2!ErledigtAm = "01.01.1999 15:00"
                        besAdoRec2!ErledigtVon = 0
                        besAdoRec2!Aconto = 0
                        besAdoRec2!Offen = 0
                        If (Val(besAdoRec2!RezeptNr) = 0) Then
                            If (AbholerDetailAdoRec!Offen < 0) Then
                                besAdoRec2!Aconto = Abs(AbholerDetailAdoRec!Offen)
                                AbholerDetailAdoRec!Offen = 0
                                AbholerDetailAdoRec!Aconto = AbholerDetailAdoRec!Preis * AbholerDetailAdoRec!lm
                            End If
                            besAdoRec2!Offen = clsOpTool.FNX(besAdoRec2!Preis * besAdoRec2!lm - besAdoRec2!Aconto)
                        End If
                        besAdoRec2.Update
'                        besRec2.Bookmark = besRec2.LastModified
                    
                        
                        lRecs = 0
                        SQLStr = "INSERT INTO AbholerInfo (Status) VALUES (" + CStr(5) + ")"
                        Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                        If (lRecs = 1) Then
                            On Error Resume Next
                            AbholerInfoAdoRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                            AbholerInfoAdoRec.Open SQLStr$, AbholerConn
                            If Not (AbholerInfoAdoRec.EOF) Then
                                NewId& = clsOpTool.CheckNullLong(AbholerInfoAdoRec!NewId)
                            End If
                        
                            AbholerInfoAdoRec.Close
                            SQLStr$ = "SELECT * FROM AbholerInfo WHERE Id=" + CStr(NewId)
                            AbholerInfoAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                        
                            AbholerInfoAdoRec!AbholerDetailId = besAdoRec2!Id ' AbholerDetailId&
                            AbholerInfoAdoRec!InfoInd = 1
                            AbholerInfoAdoRec!Status = 2
                            AbholerInfoAdoRec!AngelegtAm = Now
                            AbholerInfoAdoRec!AngelegtVon = BesorgerBenutzer%
                            AbholerInfoAdoRec!AngelegtComputer = Val(Para1.user)
                            AbholerInfoAdoRec!AngelegtLiRe = 0
                            AbholerInfoAdoRec!text = "Hinweis: angelegt durch Teil-Lieferung!"
                            
                            AbholerInfoAdoRec.Update
                        End If
                        
                        AbholerDetailAdoRec!bm = AbholerDetailAdoRec!lm
                        
                    End If
                    
                    AbholerDetailAdoRec!Status = 3
                    AbholerDetailAdoRec!ErledigtAm = Now
                    AbholerDetailAdoRec!ErledigtVon = BesorgerBenutzer%
                    If (AbholerNummerAdoRec!ProfilNr > 0) Then
                        If (MachAuftrag%) Then
                            AbholerDetailAdoRec!Status = 4
                            AbholerDetailAdoRec!AbgeholtAm = Now
                            AbholerDetailAdoRec!AbgeholtBei = BesorgerBenutzer%
                            Call SpeicherBezug
                            
                            BesorgerFertig = True
                        End If
                    End If
                    
                    If (eRezTaskId <> "") And (ChargenNr <> "") Then
                        SQLStr = "UPDATE TI_eRezepte SET"
                        SQLStr = SQLStr + " ChargenNr='" + ChargenNr + "'"
                        SQLStr = SQLStr + " WHERE TaskId='" + eRezTaskId + "'"
                        'MsgBox(SQLStr)
                        Call VerkaufDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                    End If


                End If
                If (bStatus% And 8) Then
                    AbholerDetailAdoRec!OrgPzn = AbholerDetailAdoRec!pzn
                    AbholerDetailAdoRec!OrgText = AbholerDetailAdoRec!text
                    AbholerDetailAdoRec!Orgmenge = AbholerDetailAdoRec!menge
                    AbholerDetailAdoRec!Orgeinheit = AbholerDetailAdoRec!einheit
                    AbholerDetailAdoRec!Orgpreis = AbholerDetailAdoRec!Preis
                    
                    bPzn$ = Left$(bPzn$, 8)
                    AbholerDetailAdoRec!pzn = Val(bPzn$)
                    
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(bPzn$)
                    If (TaxeAdoDBok) Then
                        On Error Resume Next
                        TaxeAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
                        If (TaxeAdoRec.EOF = False) Then
                            AbholerDetailAdoRec!text = Trim(TaxeAdoRec!Name)
                            AbholerDetailAdoRec!menge = Trim(TaxeAdoRec!menge)
                            AbholerDetailAdoRec!einheit = Trim(TaxeAdoRec!einheit)
                        Else
                            AbholerDetailAdoRec!text = "(nicht in Taxe vorhanden)"
                            AbholerDetailAdoRec!menge = ""
                            AbholerDetailAdoRec!einheit = ""
                        End If
                    Else
                        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        If (TaxeRec.EOF = False) Then
                            AbholerDetailAdoRec!text = Trim(TaxeRec!Name)
                            AbholerDetailAdoRec!menge = Trim(TaxeRec!menge)
                            AbholerDetailAdoRec!einheit = Trim(TaxeRec!einheit)
                        Else
                            AbholerDetailAdoRec!text = "(nicht in Taxe vorhanden)"
                            AbholerDetailAdoRec!menge = ""
                            AbholerDetailAdoRec!einheit = ""
                        End If
                    End If
                End If
                If (bStatus% And 16) And (bMenge% > 0) Then
                    
                    NeuInfoInd% = 1
                    SQLStr$ = "SELECT max(InfoInd) as MaxInfoInd FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailAdoRec!Id)
                    On Error Resume Next
                    besAdoRec2.Close
                    Err.Clear
                    On Error GoTo DefErr
                    besAdoRec2.Open SQLStr, AbholerConn ', adOpenDynamic, adLockOptimistic
                    If Not (besAdoRec2.EOF) Then
                        On Error Resume Next
                        NeuInfoInd% = besAdoRec2!MaxInfoInd + 1
                        On Error GoTo DefErr
                    End If
                    
                    lRecs = 0
                    SQLStr = "INSERT INTO AbholerInfo (Status) VALUES (" + CStr(5) + ")"
                    Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                    If (lRecs = 1) Then
                        On Error Resume Next
                        AbholerInfoAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                        AbholerInfoAdoRec.Open SQLStr$, AbholerConn
                        If Not (AbholerInfoAdoRec.EOF) Then
                            NewId& = clsOpTool.CheckNullLong(AbholerInfoAdoRec!NewId)
                        End If
                    
                        AbholerInfoAdoRec.Close
                        SQLStr$ = "SELECT * FROM AbholerInfo WHERE Id=" + CStr(NewId)
                        AbholerInfoAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                    
                        AbholerInfoAdoRec!AbholerDetailId = AbholerDetailAdoRec!Id ' AbholerDetailId&
                        AbholerInfoAdoRec!InfoInd = NeuInfoInd%
                        AbholerInfoAdoRec!Status = 2
                        AbholerInfoAdoRec!AngelegtAm = Now
                        AbholerInfoAdoRec!AngelegtVon = BesorgerBenutzer%
                        AbholerInfoAdoRec!AngelegtComputer = Val(Para1.user)
                        AbholerInfoAdoRec!AngelegtLiRe = 0
                        AbholerInfoAdoRec!text = "Hinweis: Teil-Lieferung!"
                        
                        AbholerInfoAdoRec.Update
                    End If
                    
                    NeuAbholerInd% = 1
                    SQLStr$ = "SELECT max(AbholerInd) as MaxAbholerInd FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
                    On Error Resume Next
                    besAdoRec2.Close
                    Err.Clear
                    On Error GoTo DefErr
                    besAdoRec2.Open SQLStr, AbholerConn ', adOpenDynamic, adLockOptimistic
                    If Not (besAdoRec2.EOF) Then
                        On Error Resume Next
                        NeuAbholerInd% = besAdoRec2!MaxAbholerInd + 1
                        On Error GoTo DefErr
                    End If
                    
                    SQLStr$ = "SELECT * INTO #tmp FROM AbholerDetails WHERE Id=" + Str$(AbholerDetailAdoRec!Id)
                    SQLStr$ = SQLStr$ + ";"
                    SQLStr = SQLStr + "ALTER TABLE #tmp DROP Column Id" + ";"
                    SQLStr = SQLStr + "INSERT INTO AbholerDetails SELECT * FROM #tmp" + ";"
                    SQLStr = SQLStr + "DROP Table #tmp" + ";"
                    Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)

                    On Error Resume Next
                    besAdoRec2.Close
                    Err.Clear
                    On Error GoTo DefErr
                    SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                    besAdoRec2.Open SQLStr$, AbholerConn
                    If Not (besAdoRec2.EOF) Then
                        NewId& = clsOpTool.CheckNullLong(besAdoRec2!NewId)
                    End If
                
                    besAdoRec2.Close
                    SQLStr$ = "SELECT * FROM AbholerDetails WHERE Id=" + CStr(NewId)
                    besAdoRec2.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                    
                    besAdoRec2!AbholerInd = NeuAbholerInd%
                    besAdoRec2!bm = bMenge
                    besAdoRec2!lm = bMenge
                    besAdoRec2!Status = 1
                    besAdoRec2!ErledigtAm = "01.01.1999 15:00"
                    besAdoRec2!ErledigtVon = 0
                    besAdoRec2!Aconto = 0
                    besAdoRec2!Offen = 0
                    If (Val(besAdoRec2!RezeptNr) = 0) Then
                        If (AbholerDetailAdoRec!Offen < 0) Then
                            besAdoRec2!Aconto = Abs(AbholerDetailAdoRec!Offen)
                            AbholerDetailAdoRec!Offen = 0
                            AbholerDetailAdoRec!Aconto = AbholerDetailAdoRec!Preis * AbholerDetailAdoRec!lm
                        End If
                        besAdoRec2!Offen = clsOpTool.FNX(besAdoRec2!Preis * besAdoRec2!lm - besAdoRec2!Aconto)
                    End If
                    besAdoRec2.Update
'                        besRec2.Bookmark = besRec2.LastModified
                
                    lRecs = 0
                    SQLStr = "INSERT INTO AbholerInfo (Status) VALUES (" + CStr(5) + ")"
                    Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                    If (lRecs = 1) Then
                        On Error Resume Next
                        AbholerInfoAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                        AbholerInfoAdoRec.Open SQLStr$, AbholerConn
                        If Not (AbholerInfoAdoRec.EOF) Then
                            NewId& = clsOpTool.CheckNullLong(AbholerInfoAdoRec!NewId)
                        End If
                    
                        AbholerInfoAdoRec.Close
                        SQLStr$ = "SELECT * FROM AbholerInfo WHERE Id=" + CStr(NewId)
                        AbholerInfoAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                    
                        AbholerInfoAdoRec!AbholerDetailId = besAdoRec2!Id ' AbholerDetailId&
                        AbholerInfoAdoRec!InfoInd = 1
                        AbholerInfoAdoRec!Status = 2
                        AbholerInfoAdoRec!AngelegtAm = Now
                        AbholerInfoAdoRec!AngelegtVon = BesorgerBenutzer%
                        AbholerInfoAdoRec!AngelegtComputer = Val(Para1.user)
                        AbholerInfoAdoRec!AngelegtLiRe = 0
                        AbholerInfoAdoRec!text = "Hinweis: angelegt durch Teil-Lieferung!"
                        
                        AbholerInfoAdoRec.Update
                    End If
                        
                End If
                
                AbholerDetailAdoRec.Update
            End If
        End If
        
        If (bStatus% And 4) Then
            KisteStatus% = 4
            BesorgerStatus% = 4
            
            SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
            SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
            On Error Resume Next
            AbholerDetailAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            AbholerDetailAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
            
            For i% = 0 To 50
                If (AbholerDetailAdoRec.EOF) Then
                    Exit For
                End If
                
                If (AbholerDetailAdoRec!AbholerTyp = 1) Then
                    If (BesorgerStatus% > 3) And (AbholerDetailAdoRec!Status < 4) Then BesorgerStatus% = 3
                    If (AbholerDetailAdoRec!Status < 3) And (BesorgerStatus% > 1) Then BesorgerStatus% = 1
                End If
                        
                If (KisteStatus% > 3) And (AbholerDetailAdoRec!Status < 4) Then KisteStatus% = 3
                If (AbholerDetailAdoRec!Status < 3) And (KisteStatus% > 1) Then KisteStatus% = 1
                
                AbholerDetailAdoRec.MoveNext
            Next i%
            
'            AbholerNummerRec.Edit
            AbholerNummerAdoRec!Status = KisteStatus%
            AbholerNummerAdoRec!StatusBesorger = BesorgerStatus%
            AbholerNummerAdoRec.Update
        End If
        
        Call clsOpTool.ErhoeheCounter
    End If

    AbholerConn.Close
ElseIf (AbholerMdb%) Then
    Set AbholerDB = OpenDatabase(ABHOLER_DB, False, False)
        
    Set AbholerParaRec = AbholerDB.OpenRecordset("Parameter", dbOpenTable)
    AbholerParaRec.Index = "Unique"
    
    SQLStr$ = "SELECT * FROM AbholerNummern WHERE AbholerNummer=" + Str$(Nummer%)
    Set AbholerNummerRec = AbholerDB.OpenRecordset(SQLStr$)
    iOk% = (AbholerNummerRec.RecordCount > 0)
    If (iOk%) Then
        kStatus% = clsOpTool.CheckNullLong(AbholerNummerRec!Status)
        iOk% = (kStatus% >= 1) And (kStatus% <= 3)
    End If
    If (iOk%) Then
        SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerRec!AbholerNummer)
        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
        Set AbholerDetailRec = AbholerDB.OpenRecordset(SQLStr$)
        
        DetailId& = -1
        For i% = 0 To 50
            If (AbholerDetailRec.EOF) Then
                Exit For
            End If
            
            If (AbholerDetailRec!pzn = Val(pzn$)) Then
                If (AbholerDetailRec!Status < 3) Then
                    DetailId& = AbholerDetailRec!Id
                    Exit For
                ElseIf (DetailId& < 0) Then
                    DetailId& = AbholerDetailRec!Id
                End If
            End If
            
            AbholerDetailRec.MoveNext
        Next i%
        
        If (DetailId& >= 0) Then
            BesorgerBenutzer% = Para1.AktBenutzer
            
            SQLStr$ = "SELECT * FROM AbholerDetails WHERE Id=" + Str$(DetailId&)
            Set AbholerDetailRec = AbholerDB.OpenRecordset(SQLStr$)
            If (AbholerDetailRec!pzn = Val(pzn$)) Then
                AbholerDetailRec.Edit
                
                If (bStatus% And 1) Then
                    AbholerDetailRec!lm = bMenge%
                    If (Val(AbholerDetailRec!RezeptNr) = 0) Then
'                        AbholerDetailRec!Offen = AbholerDetailRec!Preis * AbholerDetailRec!lm - clsOpTool.xVal(AbholerDetailRec!Aconto)
                        AbholerDetailRec!Offen = clsOpTool.FNX(AbholerDetailRec!Preis * AbholerDetailRec!lm - AbholerDetailRec!Aconto)
                    End If
                End If
                If (bStatus% And 2) Then
                    AbholerDetailRec!Preis = bPreis#
                    If (Val(AbholerDetailRec!RezeptNr) = 0) Then
'                        AbholerDetailRec!Offen = AbholerDetailRec!Preis * AbholerDetailRec!lm - clsOpTool.xVal(AbholerDetailRec!Aconto)
                        AbholerDetailRec!Offen = clsOpTool.FNX(AbholerDetailRec!Preis * AbholerDetailRec!lm - AbholerDetailRec!Aconto)
                    End If
                End If
                If (bStatus% And 4) And (AbholerDetailRec!Status <> 4) Then
                    If (AbholerDetailRec!lm < AbholerDetailRec!bm) Then
                        NeuInfoInd% = 1
                        SQLStr$ = "SELECT max(InfoInd) as MaxInfoInd FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailRec!Id)
                        Set besRec2 = AbholerDB.OpenRecordset(SQLStr$)
                        If (besRec2.RecordCount > 0) Then
                            On Error Resume Next
                            NeuInfoInd% = besRec2!MaxInfoInd + 1
                            On Error GoTo DefErr
                        End If
        
                        Set AbholerInfoRec = AbholerDB.OpenRecordset("AbholerInfo", dbOpenTable)
                        AbholerInfoRec.Index = "Id"
                        
                        AbholerInfoRec.AddNew
                        AbholerInfoRec!AbholerDetailId = AbholerDetailRec!Id ' AbholerDetailId&
                        AbholerInfoRec!InfoInd = NeuInfoInd%
                        AbholerInfoRec!Status = 2
                        AbholerInfoRec!AngelegtAm = Now
                        AbholerInfoRec!AngelegtVon = BesorgerBenutzer%
                        AbholerInfoRec!AngelegtComputer = Val(Para1.user)
                        AbholerInfoRec!AngelegtLiRe = 0
                        AbholerInfoRec!text = "Hinweis: Teil-Lieferung!"
                        AbholerInfoRec.Update
                        
                        NeuAbholerInd% = 1
'                        SQLStr$ = "SELECT max(AbholerInd) as MaxAbholerInd FROM AbholerDetails WHERE Id=" + Str$(DetailId&)
                        SQLStr$ = "SELECT max(AbholerInd) as MaxAbholerInd FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerRec!AbholerNummer)
                        Set besRec2 = AbholerDB.OpenRecordset(SQLStr$)
                        If (besRec2.RecordCount > 0) Then
                            NeuAbholerInd% = besRec2!MaxAbholerInd + 1
                        End If
                        
                        Set besRec2 = AbholerDB.OpenRecordset("AbholerDetails", dbOpenTable)
                        besRec2.Index = "Id"
                        besRec2.AddNew
                        For Each Fld In AbholerDetailRec.Fields
                            If (Fld.Name = "Id") Then
                            ElseIf (Fld.Name = "AbholerInd") Then
                                besRec2(Fld.Name) = NeuAbholerInd%
                            Else
                                besRec2(Fld.Name) = Fld.Value
                            End If
                        Next Fld
                        besRec2!bm = AbholerDetailRec!bm - AbholerDetailRec!lm
                        besRec2!lm = besRec2!bm
                        besRec2!Status = 1
                        besRec2!ErledigtAm = "01.01.1999 15:00"
                        besRec2!ErledigtVon = 0
                        besRec2!Aconto = 0
                        besRec2!Offen = 0
                        If (Val(besRec2!RezeptNr) = 0) Then
                            If (AbholerDetailRec!Offen < 0) Then
                                besRec2!Aconto = Abs(AbholerDetailRec!Offen)
                                AbholerDetailRec!Offen = 0
                                AbholerDetailRec!Aconto = AbholerDetailRec!Preis * AbholerDetailRec!lm
                            End If
                            besRec2!Offen = clsOpTool.FNX(besRec2!Preis * besRec2!lm - besRec2!Aconto)
                        End If
                        besRec2.Update
                        besRec2.Bookmark = besRec2.LastModified
                    
                        AbholerInfoRec.AddNew
                        AbholerInfoRec!AbholerDetailId = besRec2!Id ' AbholerDetailId&
                        AbholerInfoRec!InfoInd = 1
                        AbholerInfoRec!Status = 2
                        AbholerInfoRec!AngelegtAm = Now
                        AbholerInfoRec!AngelegtVon = BesorgerBenutzer%
                        AbholerInfoRec!AngelegtComputer = Val(Para1.user)
                        AbholerInfoRec!AngelegtLiRe = 0
                        AbholerInfoRec!text = "Hinweis: angelegt durch Teil-Lieferung!"
                        AbholerInfoRec.Update

                        AbholerDetailRec!bm = AbholerDetailRec!lm
                        
'                        Call WawiAddNew(wRec2)  'wRec2.AddNew
'
'                        For Each Fld In WawiRec.Fields
'                            wRec2(Fld.Name) = Fld.Value
'                        Next Fld
'
'                        If (WawiRec!bm < 0) Then SollMenge% = -SollMenge%
'                        wRec2!bm = SollMenge%
'                        wRec2!aktivlief = 0
'                        wRec2!best = " "
'                        wRec2!zugeordnet = "N"
'                        wRec2!zukontrollieren = "0" ' Chr$(0)
'                        wRec2!fixiert = "0" ' Chr$(0)
'                        If (wRec2!absage < 100) Then
'                            wRec2!absage = wRec2!absage + 1 'Lieferant%
'                        End If
'                        If (wRec2!absage = 1) Then
'                            wRec2!lief = 0
'                        End If
'                        wRec2!IstSchwellArtikel = 0
'                        wRec2!beklaufnr = CalcLaufNr&(pzn$)
'
'                        Call WawiUpdate(wRec2)  'wRec2.Update
                    End If
                    
                    AbholerDetailRec!Status = 3
                    AbholerDetailRec!ErledigtAm = Now
                    AbholerDetailRec!ErledigtVon = BesorgerBenutzer%
                    If (AbholerNummerRec!ProfilNr > 0) Then
                        If (MachAuftrag%) Then
                            AbholerDetailRec!Status = 4
                            AbholerDetailRec!AbgeholtAm = Now
                            AbholerDetailRec!AbgeholtBei = BesorgerBenutzer%
                            Call SpeicherBezug
                            
                            BesorgerFertig = True
                        End If
                    End If
                End If
                If (bStatus% And 8) Then
                    AbholerDetailRec!OrgPzn = AbholerDetailRec!pzn
                    AbholerDetailRec!OrgText = AbholerDetailRec!text
                    AbholerDetailRec!Orgmenge = AbholerDetailRec!menge
                    AbholerDetailRec!Orgeinheit = AbholerDetailRec!einheit
                    AbholerDetailRec!Orgpreis = AbholerDetailRec!Preis
                    
                    bPzn$ = Left$(bPzn$, 8)
                    AbholerDetailRec!pzn = Val(bPzn$)
                    
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(bPzn$)
                    If (TaxeAdoDBok) Then
                        On Error Resume Next
                        TaxeAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
                        If (TaxeAdoRec.EOF = False) Then
                            AbholerDetailRec!text = Trim(TaxeAdoRec!Name)
                            AbholerDetailRec!menge = Trim(TaxeAdoRec!menge)
                            AbholerDetailRec!einheit = Trim(TaxeAdoRec!einheit)
                        Else
                            AbholerDetailRec!text = "(nicht in Taxe vorhanden)"
                            AbholerDetailRec!menge = ""
                            AbholerDetailRec!einheit = ""
                        End If
                    Else
                        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        If (TaxeRec.EOF = False) Then
                            AbholerDetailRec!text = Trim(TaxeRec!Name)
                            AbholerDetailRec!menge = Trim(TaxeRec!menge)
                            AbholerDetailRec!einheit = Trim(TaxeRec!einheit)
                        Else
                            AbholerDetailRec!text = "(nicht in Taxe vorhanden)"
                            AbholerDetailRec!menge = ""
                            AbholerDetailRec!einheit = ""
                        End If
                    End If
                End If
                If (bStatus% And 16) And (bMenge% > 0) Then
                    NeuInfoInd% = 1
                    SQLStr$ = "SELECT max(InfoInd) as MaxInfoInd FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailRec!Id)
                    Set besRec2 = AbholerDB.OpenRecordset(SQLStr$)
                    If (besRec2.RecordCount > 0) Then
                        On Error Resume Next
                        NeuInfoInd% = besRec2!MaxInfoInd + 1
                        On Error GoTo DefErr
                    End If
    
                    Set AbholerInfoRec = AbholerDB.OpenRecordset("AbholerInfo", dbOpenTable)
                    AbholerInfoRec.Index = "Id"
                    
                    AbholerInfoRec.AddNew
                    AbholerInfoRec!AbholerDetailId = AbholerDetailRec!Id ' AbholerDetailId&
                    AbholerInfoRec!InfoInd = NeuInfoInd%
                    AbholerInfoRec!Status = 2
                    AbholerInfoRec!AngelegtAm = Now
                    AbholerInfoRec!AngelegtVon = BesorgerBenutzer%
                    AbholerInfoRec!AngelegtComputer = Val(Para1.user)
                    AbholerInfoRec!AngelegtLiRe = 0
                    AbholerInfoRec!text = "Hinweis: Teil-Lieferung!"
                    AbholerInfoRec.Update
                    
                    NeuAbholerInd% = 1
                    SQLStr$ = "SELECT max(AbholerInd) as MaxAbholerInd FROM AbholerDetails WHERE Id=" + Str$(DetailId&)
                    Set besRec2 = AbholerDB.OpenRecordset(SQLStr$)
                    If (besRec2.RecordCount > 0) Then
                        NeuAbholerInd% = besRec2!MaxAbholerInd + 1
                    End If
                    
                    Set besRec2 = AbholerDB.OpenRecordset("AbholerDetails", dbOpenTable)
                    besRec2.Index = "Id"
                    besRec2.AddNew
                    For Each Fld In AbholerDetailRec.Fields
                        If (Fld.Name = "Id") Then
                        ElseIf (Fld.Name = "AbholerInd") Then
                            besRec2(Fld.Name) = NeuAbholerInd%
                        Else
                            besRec2(Fld.Name) = Fld.Value
                        End If
                    Next Fld
                    besRec2!bm = bMenge%
                    besRec2!lm = bMenge%
                    besRec2!Status = 1
                    besRec2!ErledigtAm = "01.01.1999 15:00"
                    besRec2!ErledigtVon = 0
                    besRec2!Aconto = 0
                    besRec2!Offen = 0
                    If (Val(besRec2!RezeptNr) = 0) Then
                        If (AbholerDetailRec!Offen < 0) Then
                            besRec2!Aconto = Abs(AbholerDetailRec!Offen)
                            AbholerDetailRec!Offen = 0
                            AbholerDetailRec!Aconto = AbholerDetailRec!Preis * AbholerDetailRec!lm
                        End If
                        besRec2!Offen = clsOpTool.FNX(besRec2!Preis * besRec2!lm - besRec2!Aconto)
                    End If
                    besRec2.Update
                    besRec2.Bookmark = besRec2.LastModified
                
                    AbholerInfoRec.AddNew
                    AbholerInfoRec!AbholerDetailId = besRec2!Id ' AbholerDetailId&
                    AbholerInfoRec!InfoInd = 1
                    AbholerInfoRec!Status = 2
                    AbholerInfoRec!AngelegtAm = Now
                    AbholerInfoRec!AngelegtVon = BesorgerBenutzer%
                    AbholerInfoRec!AngelegtComputer = Val(Para1.user)
                    AbholerInfoRec!AngelegtLiRe = 0
                    AbholerInfoRec!text = "Hinweis: angelegt durch Teil-Lieferung!"
                    AbholerInfoRec.Update
                    
                
                End If
                
                AbholerDetailRec.Update
            End If
        End If
        
        If (bStatus% And 4) Then
            KisteStatus% = 4
            BesorgerStatus% = 4
            
            SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerRec!AbholerNummer)
            SQLStr$ = SQLStr$ + " ORDER BY AbholerInd"
            Set AbholerDetailRec = AbholerDB.OpenRecordset(SQLStr$)
            
            For i% = 0 To 50
                If (AbholerDetailRec.EOF) Then
                    Exit For
                End If
                
                If (AbholerDetailRec!AbholerTyp = 1) Then
                    If (BesorgerStatus% > 3) And (AbholerDetailRec!Status < 4) Then BesorgerStatus% = 3
                    If (AbholerDetailRec!Status < 3) And (BesorgerStatus% > 1) Then BesorgerStatus% = 1
                End If
                        
                If (KisteStatus% > 3) And (AbholerDetailRec!Status < 4) Then KisteStatus% = 3
                If (AbholerDetailRec!Status < 3) And (KisteStatus% > 1) Then KisteStatus% = 1
                
                AbholerDetailRec.MoveNext
            Next i%
            
            AbholerNummerRec.Edit
            AbholerNummerRec!Status = KisteStatus%
            AbholerNummerRec!StatusBesorger = BesorgerStatus%
            AbholerNummerRec.Update
        End If
        
        Call clsOpTool.ErhoeheCounter
    End If
        
    AbholerDB.Close
Else
    If (Abs(Nummer%) > 999) Then
        Call clsError.DefErrPop: Exit Function
    End If
    
    Set Kiste1 = New clsKiste
    Kiste1.OpenDatei
    Belegt% = Kiste1.Belegt(Nummer%)
    If (Belegt%) Then
        Call Kiste1.GetKiste(Nummer%)
            
        kInd% = -1
        For i% = 0 To 9
            Kiste1.GetInhalt (i%)
            If (Kiste1.pzn = pzn$) Then
                If (Kiste1.Status < 3) Then
                    kInd% = i%
                    Exit For
                ElseIf (kInd% < 0) Then
                    kInd% = i%
                End If
            End If
        Next i%
        
        If (kInd% >= 0) Then
            i% = kInd%
    '    For i% = 0 To 9
            Kiste1.GetInhalt (i%)
            If (Kiste1.pzn = pzn$) Then
                If (bStatus% And 1) Then
                    Kiste1.menge = bMenge%
                End If
                If (bStatus% And 2) Then
                    Kiste1.Preis = bPreis#
                End If
                If (bStatus% And 4) And (Kiste1.Status <> 4) Then
                    Kiste1.Status = 3
                    If (Kiste1.ProfilNr > 0) Then
                        If (Kiste1.MachAuftrag) Then
                            Kiste1.Status = 4
                        End If
                    End If
                End If
                If (bStatus% And 8) Then
                    Kiste1.pzn = bPzn$
                End If
                Kiste1.PutInhalt (i%)
    '            Exit For
            End If
    '    Next i%
        End If
        
        If (bStatus% And 4) Then
            KisteStatus% = 4
            BesorgerStatus% = 4
      
            For i% = 0 To 9
                erg% = Kiste1.GetInhalt(i%)
                If (erg%) Then
                    If (Kiste1.WasTun = "B") Then
                        If (BesorgerStatus% > 3) And (Kiste1.Status < 4) Then BesorgerStatus% = 3
                        If (Kiste1.Status < 3) And (BesorgerStatus% > 1) Then BesorgerStatus% = 1
                    End If
                    If (InStr("RTBHP", Kiste1.WasTun) > 0) Then
                        If (KisteStatus% > 3) And (Kiste1.Status < 4) Then KisteStatus% = 3
                        If (Kiste1.Status < 3) And (KisteStatus% > 1) Then KisteStatus% = 1
                    End If
                End If
            Next i%
      
            Arbeit$ = Kiste1.Arbeit
            Mid$(Arbeit$, 4, 1) = Chr$(BesorgerStatus%)
            Kiste1.Arbeit = Arbeit$
            Kiste1.KistenStatus = KisteStatus%
            Kiste1.PutKiste (Nummer%)
            
      
    '        If (KisteStatus% = 4) Then
    '            '* alles abgeholt
    '            '* wenn das nur Besorger waren, dann entfernen
    '            NurFertigeBesorger% = True
    '            For i% = 1 To 5
    '                s% = Asc(Mid$(Arbeit$, i%, 1))
    '                If (i% <> 4) And (s% <> 0) Then
    '                    NurFertigeBesorger% = 0
    '                End If
    '            Next i%
    '            If (NurFertigeBesorger%) Then
    '                '* AnfBlocks ausleeren
    '                For i% = 0 To 9
    '                    erg% = Kiste1.ClearInhalt(i%)
    '                Next i%
    '
    '                Kiste1.ClearKiste (Nummer%)
    '            End If
    '        End If
        End If
    End If
        
    Kiste1.CloseDatei
    Set Kiste1 = Nothing
End If

Call clsError.DefErrPop
End Function

Function MachAuftrag%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachAuftrag%")
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
Dim clsKunden1 As New clsKunden
Dim clsKundenZus1 As New clsKundZusatz
Dim AufDB As Database
Dim ProfDB As Database
Dim kDB As Database
Dim auf As Recordset
Dim auf2 As Recordset
Dim prof As Recordset
Dim arch As Recordset
Dim kRec As Recordset
Dim kAdrRec As Recordset
Dim kAdoRec As New ADODB.Recordset
Dim kAdrAdoRec As New ADODB.Recordset
Dim i%, j%, KZUSTELL%, KZUST_IDX%, zahlart%, ok%, IstMischKunde%, KundeGebFrei%, AuftragAdd%, HatLieferSchein%, PreisSchonBerechnet As Boolean
Dim dBesAconto#, dBesPreis#, dBesOffen#
Dim recno&, LieferScheinNr&
Dim s$, AufNr$, ZustName$(4), Zeitraum$, SQLStr$, LSnr$, BesPzn$
Dim tRec2 As Recordset
Dim AbholerAuftragsNr$
Dim KundenDB1 As clsKundenDB

If (AbholerSQL) Then
    MachAuftrag = MachAuftragAdo
    Call clsError.DefErrPop: Exit Function
End If

MachAuftrag% = 0

On Error Resume Next
Err.Clear
Set ProfDB = OpenDatabase("PROFILE.MDB")
Set prof = ProfDB.OpenRecordset("Profile", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    Call clsError.DefErrPop: Exit Function
End If

prof.Index = "Unique"
prof.Seek "=", AbholerNummerRec!ProfilNr
ok% = 0
If (prof.NoMatch = False) Then
    If (prof!AbhFertig) Then
        ok% = True
    End If
End If
If (ok% = 0) Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If

        
'dBesAconto# = clsOpTool.FNX(clsOpTool.xVal(BesAconto$))
'dBesPreis# = clsOpTool.FNX(BesMenge% * clsOpTool.xVal(aw.Preis))
dBesAconto# = AbholerDetailRec!Aconto
dBesOffen# = AbholerDetailRec!Offen
dBesPreis# = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis)
BesPzn$ = clsOpTool.PznString(AbholerDetailRec!pzn)

'LieferscheinNr& = 0
s$ = ""
On Error Resume Next
s$ = Trim(AbholerNummerRec!LieferScheinNr)
HatLieferSchein% = (s$ <> "")
If HatLieferSchein% <> 0 Then
  PreisSchonBerechnet = True
  i% = 0
  i% = AbholerNummerRec!AuftragsLS
  If i% <> 0 Then PreisSchonBerechnet = False
End If
'If (HatLieferschein%) Then
'    LieferscheinNr& = Val(s$)
'End If

Err.Clear
If Not IsNull(AbholerNummerRec!Auftragsnummer) Then
  AbholerAuftragsNr$ = AbholerNummerRec!Auftragsnummer
End If
If Err.Number <> 0 Then AbholerAuftragsNr$ = ""
On Error GoTo DefErr

ok% = True
If (Val(AbholerDetailRec!RezeptNr) = 0) Then    '
    'Preisunterschied kann dem Kunden nur verrechnet werden, wenn kein Lieferschein und er jetzt einen Auftrag bekommt
    If (dBesPreis# <> dBesAconto#) And ((prof!AutoZustellung = 0) Or (HatLieferSchein%)) Then
        ok% = 0
    End If
Else
    'Wenn RezGebühr noch nicht bezahlt ist und bei der Bestellung kein Lieferschien erstellt wurde,
    'kann sie (die RezGebühr)nur im jetzt entstehenden Auftrag verrechnet werden
    If (dBesOffen# > 0#) And (prof!AutoZustellung = 0) And (HatLieferSchein% = 0) Then
        ok% = 0
    End If
End If

If (ok) And (InStr(Para1.Benutz, "j") > 0) Then
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(AbholerDetailRec!pzn)
    If (TaxeAdoDBok) Then
        On Error Resume Next
        TaxeAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
        If (TaxeAdoRec.EOF = False) Then
            ok = (TaxeAdoRec!BTMKz = 0)
        End If
    End If
End If


If (ok% = 0) Then
    Call clsDialog.MessageBox("Abholer " + Format(AbholerNummerRec!AbholerNummer, "0") + " muss an der Kasse fertiggemacht werden!", vbInformation)
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If


Err.Clear
Set AufDB = OpenDatabase("AUFTRAG.MDB")
Set auf = AufDB.OpenRecordset("Aufträge", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If
 

If (Para1.Land = "D") Or (Dir("kunbezug.mdb") <> "") Then
    MachAuftrag% = True
    
    Dim AuftragMitBesorger%
    AuftragMitBesorger = 0
    s$ = Space(10)
    i% = GetPrivateProfileString("Securpharm", "AuftragMitBesorger", "0", s$, 255, CurDir + "\phcare.ini")
    If (i% > 0) Then
        AuftragMitBesorger = Val(Left(s$, i%))
    End If
    If (AuftragMitBesorger = 1) Then
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If


    If (prof!AutoZustellung = 0) Then
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If
    
    
    Set KundenDB1 = New clsKundenDB
    KundenDBok = KundenDB1.DBvorhanden
    If (KundenDBok) Then
        KundenDBok = KundenDB1.OpenDB
    End If
    If (KundenDBok) Then
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(AbholerNummerRec!KuNr)
        kAdoRec.Open SQLStr$, KundenDB1.ActiveConn
        If (kAdoRec.EOF) Then
            AufDB.Close
            ProfDB.Close
            kDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        
        IstMischKunde% = kAdoRec!MischKunde
        
        KundeGebFrei% = 0
        If Not IsNull(kAdoRec!BefreiungBis) And Not IsEmpty(kAdoRec!BefreiungBis) Then
            If (kAdoRec!BefreiungBis >= Now) Then
                KundeGebFrei% = True
            End If
        End If
        
        AuftragAdd% = 0
        If (IstMischKunde% = 0) Then
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
            SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
            SQLStr$ = SQLStr$ + " AND FERTIG=0"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If (auf2.EOF = False) Then
                auf2.MoveFirst
                Do
                    If (auf2.EOF) Then
                        Exit Do
                    End If
                    
                    If (DateDiff("s", auf2!datum, Now) < 60) Then
                        AuftragAdd% = True
                        Exit Do
                    End If
                    
                    auf2.MoveNext
                Loop
            End If
        ElseIf AbholerAuftragsNr$ > "" Then
            On Error Resume Next
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
            SQLStr$ = SQLStr$ + " AND Auftragsnr = '" + AbholerAuftragsNr$ + "'"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If Err.Number = 0 Then
              If (auf2.EOF = False) Then
                  auf2.MoveFirst
                  AuftragAdd% = True
              End If
            End If
            On Error GoTo DefErr
        End If
            
        If (AuftragAdd%) Then
            auf2.Edit
                
            If (Val(AbholerDetailRec!RezeptNr) > 0) Then    '
                auf2!RezAnz = 1
                
                If (HatLieferSchein%) Then
        '            If (LieferscheinNr& > 0) Then
                    If Not (PreisSchonBerechnet) Then
                        auf2!Summe = auf2!Summe + dBesAconto#
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                        Else
    '                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
    '                        End If
    '                    End If
                    End If
        '            End If
                Else
                    auf2!Summe = auf2!Summe + dBesOffen#
                End If
            Else
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
        '            If (LieferscheinNr& > 0) Then
                        auf2!Summe = auf2!Summe + AbholerDetailRec!Preis
        '            Else
        '            End If
                    End If
                Else
                    auf2!Summe = auf2!Summe + AbholerDetailRec!Preis - dBesAconto#
                End If
            End If
            
            auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
            
            Err.Clear
            auf2.Update
        Else
            auf.Index = "Unique"
            AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
            s$ = Space(15)
            i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
            If (i% > 0) Then
                s$ = Left(s$, i%)
                If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
                    AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
                End If
            End If
                
            Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
            Do
                i% = 0
                Call clsOpTool.EanPruef(AufNr$)
                auf.Seek "=", AufNr$
                If Not (auf.NoMatch) Then i% = 1
                arch.Index = "Unique"
                arch.Seek "=", AufNr$
                If Not (arch.NoMatch) Then i% = 1
                If (i% = 1) Then
                    AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
                End If
            Loop While (i% = 1)
            i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
            
            Zeitraum$ = ""
            zahlart% = 1
            If (kAdoRec!SammelQuittung > 0) Then
                zahlart = 2
            End If
            
            If (clsOpTool.CheckNullLong(kAdoRec!LieferadresseKz) <> 0) Then
                SQLStr$ = "SELECT * FROM AdressenZusatz WHERE KundenNr =" + Str$(AbholerNummerRec!KuNr)
                SQLStr$ = SQLStr$ + " AND Typ=0"
                kAdrAdoRec.Open SQLStr, KundenDB1.ActiveConn
                If (kAdrAdoRec.EOF = False) Then
                    ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Titel)))
                    ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Name)))
                    ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Station)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!HausNr)))
                    ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Ort)))
                End If
                kAdrAdoRec.Close
            End If
            If (ZustName$(0) = "") And (ZustName$(1) = "") Then
                ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Titel)))
                ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Name)))
            End If
            If (ZustName$(2) = "") And (ZustName$(3) = "") Then
                ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!HausNr)))
                ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Ort)))
            End If
            If (IstMischKunde% = 0) Then
                ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!MobilPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!FonPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!FonDienst))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!MobilDienst))
            End If
            
            auf.AddNew
            auf!AuftragsNr = AufNr$
            auf!ProfilNr = AbholerNummerRec!ProfilNr
            auf!knr = AbholerNummerRec!KuNr
            auf!datum = Now
            
            LSnr$ = ""
            If (HatLieferSchein%) Then
              If Not (PreisSchonBerechnet) Then
        '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
                LSnr$ = Trim(AbholerNummerRec!LieferScheinNr)
              End If
            End If
            auf!LSnr = LSnr$
                
            auf!Zahlungsart = zahlart%
            auf!Summe = 0
            If (Val(AbholerDetailRec!RezeptNr) > 0) Then
                auf!RezAnz = 1
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
                        auf!Summe = dBesAconto#
        '            If (LieferscheinNr& > 0) Then
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                            auf!Summe = 0
    '                        Else
    '                            auf!Summe = tRec2!ZuzaNeu / 100
    '                        End If
    '                    Else
    '                        auf!Summe = 0
    '                    End If
        '            Else
        '                auf!Summe = 0
        '            End If
                    End If
                Else
                    auf!Summe = dBesOffen#
                End If
            Else
                auf!RezAnz = 0
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
    '                    auf!Summe = AbholerDetailRec!Preis
                        auf!Summe = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis)
                    End If
                Else
    '                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
                    auf!Summe = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis) - dBesAconto#
                End If
            End If
            
            prof.Index = "Unique"
            prof.Seek "=", auf!ProfilNr
            If (prof.NoMatch) Then
                auf!RezBringen = prof!RezBringen
            Else
                auf!RezBringen = 0
            End If
            
    '        auf!Positionen = 1  'ZustellAnz
            auf!Positionen = AbholerDetailRec!lm  'ZustellAnz
            auf!Kundenichtda = False
            auf!kZu1 = Left$(ZustName(0), 30)
            auf!kZu2 = Left$(ZustName(1), 30)
            auf!kZu3 = Left$(ZustName(2), 30)
            auf!kZu4 = Left$(ZustName(3), 30)
            auf!kZu5 = Left$(ZustName(4), 30)
            auf!Zeitraum = Zeitraum$
            
            Err.Clear
            auf.Update
            
        End If
        
        kAdoRec.Close
        KundenDB1.CloseDB
    Else
        On Error Resume Next
        Err.Clear
        Set kDB = OpenDatabase("kunden.mdb", False, True)
        If (Err.Number <> 0) Then
            On Error GoTo DefErr
            AufDB.Close
            ProfDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        On Error GoTo DefErr
        
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(AbholerNummerRec!KuNr)
        Set kRec = kDB.OpenRecordset(SQLStr$)
        If (kRec.EOF) Then
            AufDB.Close
            ProfDB.Close
            kDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        
        IstMischKunde% = kRec!MischKunde
        
        KundeGebFrei% = 0
        If Not IsNull(kRec!BefreiungBis) And Not IsEmpty(kRec!BefreiungBis) Then
            If (kRec!BefreiungBis >= Now) Then
                KundeGebFrei% = True
            End If
        End If
        
        AuftragAdd% = 0
        If (IstMischKunde% = 0) Then
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
            SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
            SQLStr$ = SQLStr$ + " AND FERTIG=0"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If (auf2.EOF = False) Then
                auf2.MoveFirst
                Do
                    If (auf2.EOF) Then
                        Exit Do
                    End If
                    
                    If (DateDiff("s", auf2!datum, Now) < 60) Then
                        AuftragAdd% = True
                        Exit Do
                    End If
                    
                    auf2.MoveNext
                Loop
            End If
        ElseIf AbholerAuftragsNr$ > "" Then
            On Error Resume Next
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
            SQLStr$ = SQLStr$ + " AND Auftragsnr = '" + AbholerAuftragsNr$ + "'"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If Err.Number = 0 Then
              If (auf2.EOF = False) Then
                  auf2.MoveFirst
                  AuftragAdd% = True
              End If
            End If
            On Error GoTo DefErr
        End If
            
        If (AuftragAdd%) Then
            auf2.Edit
                
            If (Val(AbholerDetailRec!RezeptNr) > 0) Then    '
                auf2!RezAnz = 1
                
                If (HatLieferSchein%) Then
        '            If (LieferscheinNr& > 0) Then
                    If Not (PreisSchonBerechnet) Then
                        auf2!Summe = auf2!Summe + dBesAconto#
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                        Else
    '                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
    '                        End If
    '                    End If
                    End If
        '            End If
                Else
                    auf2!Summe = auf2!Summe + dBesOffen#
                End If
            Else
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
        '            If (LieferscheinNr& > 0) Then
                        auf2!Summe = auf2!Summe + AbholerDetailRec!Preis
        '            Else
        '            End If
                    End If
                Else
                    auf2!Summe = auf2!Summe + AbholerDetailRec!Preis - dBesAconto#
                End If
            End If
            
            auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
            
            Err.Clear
            auf2.Update
        Else
            auf.Index = "Unique"
            AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
            s$ = Space(15)
            i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
            If (i% > 0) Then
                s$ = Left(s$, i%)
                If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
                    AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
                End If
            End If
                
            Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
            Do
                i% = 0
                Call clsOpTool.EanPruef(AufNr$)
                auf.Seek "=", AufNr$
                If Not (auf.NoMatch) Then i% = 1
                arch.Index = "Unique"
                arch.Seek "=", AufNr$
                If Not (arch.NoMatch) Then i% = 1
                If (i% = 1) Then
                    AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
                End If
            Loop While (i% = 1)
            i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
            
            Zeitraum$ = ""
            zahlart% = 1
            If (kRec!SammelQuittung > 0) Then
                zahlart = 2
            End If
            
            If (clsOpTool.CheckNullLong(kRec!LieferadresseKz) <> 0) Then
                SQLStr$ = "SELECT * FROM AdressenZusatz WHERE KundenNr =" + Str$(AbholerNummerRec!KuNr)
                SQLStr$ = SQLStr$ + " AND Typ=0"
                Set kAdrRec = kDB.OpenRecordset(SQLStr$)
                If (kAdrRec.EOF = False) Then
                    ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Titel)))
                    ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Name)))
                    ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Station)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!HausNr)))
                    ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Ort)))
                End If
            End If
            If (ZustName$(0) = "") And (ZustName$(1) = "") Then
                ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Titel)))
                ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Name)))
            End If
            If (ZustName$(2) = "") And (ZustName$(3) = "") Then
                ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kRec!HausNr)))
                ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Ort)))
            End If
            If (IstMischKunde% = 0) Then
                ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonDienst))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilDienst))
            End If
            
            auf.AddNew
            auf!AuftragsNr = AufNr$
            auf!ProfilNr = AbholerNummerRec!ProfilNr
            auf!knr = AbholerNummerRec!KuNr
            auf!datum = Now
            
            LSnr$ = ""
            If (HatLieferSchein%) Then
              If Not (PreisSchonBerechnet) Then
        '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
                LSnr$ = Trim(AbholerNummerRec!LieferScheinNr)
              End If
            End If
            auf!LSnr = LSnr$
                
            auf!Zahlungsart = zahlart%
            auf!Summe = 0
            If (Val(AbholerDetailRec!RezeptNr) > 0) Then
                auf!RezAnz = 1
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
                        auf!Summe = dBesAconto#
        '            If (LieferscheinNr& > 0) Then
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                            auf!Summe = 0
    '                        Else
    '                            auf!Summe = tRec2!ZuzaNeu / 100
    '                        End If
    '                    Else
    '                        auf!Summe = 0
    '                    End If
        '            Else
        '                auf!Summe = 0
        '            End If
                    End If
                Else
                    auf!Summe = dBesOffen#
                End If
            Else
                auf!RezAnz = 0
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
    '                    auf!Summe = AbholerDetailRec!Preis
                        auf!Summe = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis)
                    End If
                Else
    '                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
                    auf!Summe = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis) - dBesAconto#
                End If
            End If
            
            prof.Index = "Unique"
            prof.Seek "=", auf!ProfilNr
            If (prof.NoMatch) Then
                auf!RezBringen = prof!RezBringen
            Else
                auf!RezBringen = 0
            End If
            
    '        auf!Positionen = 1  'ZustellAnz
            auf!Positionen = AbholerDetailRec!lm  'ZustellAnz
            auf!Kundenichtda = False
            auf!kZu1 = Left$(ZustName(0), 30)
            auf!kZu2 = Left$(ZustName(1), 30)
            auf!kZu3 = Left$(ZustName(2), 30)
            auf!kZu4 = Left$(ZustName(3), 30)
            auf!kZu5 = Left$(ZustName(4), 30)
            auf!Zeitraum = Zeitraum$
            
            Err.Clear
            auf.Update
            
            kDB.Close
        End If
    End If
    
Else
    On Error GoTo DefErr
    
    KZUSTELL% = clsDat.FileOpen("kzustell.dat", "R", "R", 256)
    If (KZUSTELL% = 0) Then
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If
     
    FabsErrf% = clsDat.IndexOpen("kzustell.ix1", KZUST_IDX%)
    If (FabsErrf%) Then
        Close #KZUSTELL%
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If
    
    MachAuftrag% = True
    
    If (prof!AutoZustellung = 0) Then
        Close #KZUSTELL%
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If
    
    clsKundenZus1.OpenDatei
    clsKundenZus1.GetRecord (AbholerNummerRec!KuNr + 1)
    IstMischKunde% = clsKundenZus1.MischKunde
    clsKundenZus1.CloseDatei
    Set clsKundenZus1 = Nothing
    
    AuftragAdd% = 0
    If (IstMischKunde% = 0) Then
        SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
        SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
        SQLStr$ = SQLStr$ + " AND FERTIG=0"
        Set auf2 = AufDB.OpenRecordset(SQLStr$)
        If (auf2.EOF = False) Then
            auf2.MoveFirst
            Do
                If (auf2.EOF) Then
                    Exit Do
                End If
                
                If (DateDiff("s", auf2!datum, Now) < 60) Then
                    AuftragAdd% = True
                    Exit Do
                End If
                
                auf2.MoveNext
            Loop
        End If
    End If
        
    If (AuftragAdd%) Then
        Close #KZUSTELL%
        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
        
        auf2.Edit
            
        If (Val(AbholerDetailRec!RezeptNr) > 0) Then    '
            auf2!RezAnz = 1
            If (HatLieferSchein%) Then
                If Not (PreisSchonBerechnet) Then
                    auf2!Summe = auf2!Summe + dBesAconto#
    '            If (LieferscheinNr& > 0) Then
'                    If (Val(BesPzn$) > 0) Then
'                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
'                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
'                        If (tRec2.EOF) Then
'                        Else
'                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
'                        End If
'                    End If
    '            End If
                End If
            Else
                auf2!Summe = auf2!Summe + dBesOffen#
            End If
        Else
            If (HatLieferSchein%) Then
                If Not (PreisSchonBerechnet) Then
    '            If (LieferscheinNr& > 0) Then
                    auf2!Summe = auf2!Summe + AbholerDetailRec!Preis
    '            Else
    '            End If
                End If
            Else
                auf2!Summe = auf2!Summe + AbholerDetailRec!Preis - dBesAconto#
            End If
        End If
        
        auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
        
        Err.Clear
        auf2.Update
    Else
        auf.Index = "Unique"
        AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
        s$ = Space(15)
        i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
        If (i% > 0) Then
            s$ = Left(s$, i%)
            If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
                AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
            End If
        End If
            
        Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
        Do
            i% = 0
            Call clsOpTool.EanPruef(AufNr$)
            auf.Seek "=", AufNr$
            If Not (auf.NoMatch) Then i% = 1
            arch.Index = "Unique"
            arch.Seek "=", AufNr$
            If Not (arch.NoMatch) Then i% = 1
            If (i% = 1) Then
                AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
            End If
        Loop While (i% = 1)
        i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
            
        zahlart% = 1
        FabsErrf% = clsDat.IndexSearch%(KZUST_IDX%, Format(AbholerNummerRec!KuNr, "00000"), recno&)
        If (recno& > 0) Then
            Get #KZUSTELL%, recno& + 1, kzu
            
            zahlart% = Val(kzu.Zahlungsart)
            ZustName$(0) = Trim(kzu.Name)
            ZustName$(1) = Trim(kzu.zHdn)
            ZustName$(2) = Trim(kzu.Strasse)
            ZustName$(3) = Trim(kzu.PlzOrt)
            ZustName$(4) = Trim(kzu.telefon)
            Zeitraum$ = Trim(kzu.Zeitraum)
        End If
        
        Close #KZUSTELL%
        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
        
        
        clsKunden1.OpenDatei
        clsKunden1.GetRecord (AbholerNummerRec!KuNr + 1)
        
        If (ZustName$(0) = "") And (ZustName$(1) = "") Then
            ZustName$(0) = RTrim$(clsKunden1.Wert("name", 0))
            ZustName$(1) = RTrim$(clsKunden1.Wert("name", 1))
        End If
        If (ZustName$(2) = "") And (ZustName$(3) = "") Then
            ZustName$(2) = RTrim$(clsKunden1.Wert("name", 2))
            ZustName$(3) = RTrim$(clsKunden1.Wert("name", 3))
        End If
        'If (zuName(4) = "") Then
        '    GetDataRecord dbKZus, AktKundensatz
        '    zuName(4) = Trim(kzus.Tel)
        'End If
        For i% = 0 To 3
            Call OemToChar(ZustName$(i%), ZustName$(i%))
        Next i%
        Call OemToChar(Zeitraum$, Zeitraum$)
            
        auf.AddNew
        auf!AuftragsNr = AufNr$
        auf!ProfilNr = AbholerNummerRec!ProfilNr
        auf!knr = AbholerNummerRec!KuNr
        auf!datum = Now
        
        LSnr$ = ""
        If (HatLieferSchein%) Then
    '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
          If Not (PreisSchonBerechnet) Then
            LSnr$ = Trim(AbholerNummerRec!LieferScheinNr)
          End If
        End If
        auf!LSnr = LSnr$
            
        auf!Zahlungsart = zahlart%
        auf!Summe = 0
        If (Val(AbholerDetailRec!RezeptNr) > 0) Then
            auf!RezAnz = 1
            If (HatLieferSchein%) Then
                If Not (PreisSchonBerechnet) Then
                    auf!Summe = dBesAconto#
    '            If (LieferscheinNr& > 0) Then
'                    If (Val(BesPzn$) > 0) Then
'                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
'                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
'                        If (tRec2.EOF) Then
'                            auf!Summe = 0
'                        Else
'                            auf!Summe = tRec2!ZuzaNeu / 100
'                        End If
'                    Else
'                        auf!Summe = 0
'                    End If
    '            Else
    '                auf!Summe = 0
    '            End If
                End If
            Else
                auf!Summe = dBesOffen#
            End If
        Else
            auf!RezAnz = 0
            If (HatLieferSchein%) Then
                If Not (PreisSchonBerechnet) Then
    '            If (LieferscheinNr& > 0) Then
                    auf!Summe = AbholerDetailRec!Preis
    '            Else
    '                auf!Summe = 0
    '            End If
                End If
            Else
                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
            End If
        End If
        
        prof.Index = "Unique"
        prof.Seek "=", auf!ProfilNr
        If (prof.NoMatch) Then
            auf!RezBringen = prof!RezBringen
        Else
            auf!RezBringen = 0
        End If
        
        auf!Positionen = 1  'ZustellAnz
        auf!Kundenichtda = False
        auf!kZu1 = ZustName(0)
        auf!kZu2 = ZustName(1)
        auf!kZu3 = ZustName(2)
        auf!kZu4 = ZustName(3)
        auf!kZu5 = ZustName(4)
        auf!Zeitraum = Zeitraum$
        
        Err.Clear
        auf.Update
    
        clsKunden1.CloseDatei
        Set clsKunden1 = Nothing
    End If
End If
    
AufDB.Close
ProfDB.Close


s$ = Space(10)
i% = GetPrivateProfileString("LIFMANAG", "Aktualisieren", "", s$, 255, "winop.ini")
If (i% > 0) Then
    i% = Val(Left(s$, i%))
Else
    i% = 0
End If
If (i% >= 32000) Then
    i% = 0
End If
i% = i% + 1
s$ = Format(i%, "0")
j% = WritePrivateProfileString("LIFMANAG", "Aktualisieren", s$, "winop.ini")

Call clsError.DefErrPop
End Function

'für Abholer als SQL
Function MachAuftragAdo%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachAuftragAdo%")
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
'Dim clsKunden1 As New clsKunden
'Dim clsKundenZus1 As New clsKundZusatz
Dim AufDB As Database
Dim ProfDB As Database
Dim kDB As Database
Dim auf As Recordset
Dim auf2 As Recordset
Dim prof As Recordset
Dim arch As Recordset
Dim kRec As Recordset
Dim kAdrRec As Recordset
Dim kAdoRec As New ADODB.Recordset
Dim kAdrAdoRec As New ADODB.Recordset
Dim i%, j%, KZUSTELL%, KZUST_IDX%, zahlart%, ok%, IstMischKunde%, KundeGebFrei%, AuftragAdd%, HatLieferSchein%, PreisSchonBerechnet As Boolean
Dim dBesAconto#, dBesPreis#, dBesOffen#
Dim recno&, LieferScheinNr&
Dim s$, AufNr$, ZustName$(4), Zeitraum$, SQLStr$, LSnr$, BesPzn$
Dim tRec2 As Recordset
Dim AbholerAuftragsNr$
Dim KundenDB1 As clsKundenDB

MachAuftragAdo% = 0

On Error Resume Next
Err.Clear
Set ProfDB = OpenDatabase("PROFILE.MDB")
Set prof = ProfDB.OpenRecordset("Profile", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    Call clsError.DefErrPop: Exit Function
End If

prof.Index = "Unique"
prof.Seek "=", AbholerNummerAdoRec!ProfilNr
ok% = 0
If (prof.NoMatch = False) Then
    If (prof!AbhFertig) Then
        ok% = True
    End If
End If
If (ok% = 0) Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If

        
'dBesAconto# = clsOpTool.FNX(clsOpTool.xVal(BesAconto$))
'dBesPreis# = clsOpTool.FNX(BesMenge% * clsOpTool.xVal(aw.Preis))
dBesAconto# = AbholerDetailAdoRec!Aconto
dBesOffen# = AbholerDetailAdoRec!Offen
dBesPreis# = clsOpTool.FNX(AbholerDetailAdoRec!lm * AbholerDetailAdoRec!Preis)
BesPzn$ = clsOpTool.PznString(AbholerDetailAdoRec!pzn)

'LieferscheinNr& = 0
s$ = ""
On Error Resume Next
s$ = Trim(AbholerNummerAdoRec!LieferScheinNr)
HatLieferSchein% = (s$ <> "")
If HatLieferSchein% <> 0 Then
  PreisSchonBerechnet = True
  i% = 0
  i% = AbholerNummerAdoRec!AuftragsLS
  If i% <> 0 Then PreisSchonBerechnet = False
End If
'If (HatLieferschein%) Then
'    LieferscheinNr& = Val(s$)
'End If

Err.Clear
If Not IsNull(AbholerNummerAdoRec!Auftragsnummer) Then
  AbholerAuftragsNr$ = AbholerNummerAdoRec!Auftragsnummer
End If
If Err.Number <> 0 Then AbholerAuftragsNr$ = ""
On Error GoTo DefErr

ok% = True
If (Val(AbholerDetailAdoRec!RezeptNr) = 0) Then    '
    'Preisunterschied kann dem Kunden nur verrechnet werden, wenn kein Lieferschein und er jetzt einen Auftrag bekommt
    If (dBesPreis# <> dBesAconto#) And ((prof!AutoZustellung = 0) Or (HatLieferSchein%)) Then
        ok% = 0
    End If
Else
    'Wenn RezGebühr noch nicht bezahlt ist und bei der Bestellung kein Lieferschien erstellt wurde,
    'kann sie (die RezGebühr)nur im jetzt entstehenden Auftrag verrechnet werden
    If (dBesOffen# > 0#) And (prof!AutoZustellung = 0) And (HatLieferSchein% = 0) Then
        ok% = 0
    End If
End If

If (ok) And (InStr(Para1.Benutz, "j") > 0) Then
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(AbholerDetailAdoRec!pzn)
    If (TaxeAdoDBok) Then
        On Error Resume Next
        TaxeAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
        If (TaxeAdoRec.EOF = False) Then
            ok = (TaxeAdoRec!BTMKz = 0)
        End If
    End If
End If


If (ok% = 0) Then
    Call clsDialog.MessageBox("Abholer " + Format(AbholerNummerAdoRec!AbholerNummer, "0") + " muss an der Kasse fertiggemacht werden!", vbInformation)
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If


Err.Clear
Set AufDB = OpenDatabase("AUFTRAG.MDB")
Set auf = AufDB.OpenRecordset("Aufträge", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If
 

If (Para1.Land = "D") Or (Dir("kunbezug.mdb") <> "") Then
    MachAuftragAdo% = True
    
    Dim AuftragMitBesorger%
    AuftragMitBesorger = 0
    s$ = Space(10)
    i% = GetPrivateProfileString("Securpharm", "AuftragMitBesorger", "0", s$, 255, CurDir + "\phcare.ini")
    If (i% > 0) Then
        AuftragMitBesorger = Val(Left(s$, i%))
    End If
    If (AuftragMitBesorger = 1) Then
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If

    If (prof!AutoZustellung = 0) Then
        AufDB.Close
        ProfDB.Close
        Call clsError.DefErrPop: Exit Function
    End If
    
    
    Set KundenDB1 = New clsKundenDB
    KundenDBok = KundenDB1.DBvorhanden
    If (KundenDBok) Then
        KundenDBok = KundenDB1.OpenDB
    End If
    If (KundenDBok) Then
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(AbholerNummerAdoRec!KuNr)
        kAdoRec.Open SQLStr$, KundenDB1.ActiveConn
        If (kAdoRec.EOF) Then
            AufDB.Close
            ProfDB.Close
            kDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        
        IstMischKunde% = kAdoRec!MischKunde
        
        KundeGebFrei% = 0
        If Not IsNull(kAdoRec!BefreiungBis) And Not IsEmpty(kAdoRec!BefreiungBis) Then
            If (kAdoRec!BefreiungBis >= Now) Then
                KundeGebFrei% = True
            End If
        End If
        
        AuftragAdd% = 0
        If (IstMischKunde% = 0) Then
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerAdoRec!KuNr)
            SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
            SQLStr$ = SQLStr$ + " AND FERTIG=0"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If (auf2.EOF = False) Then
                auf2.MoveFirst
                Do
                    If (auf2.EOF) Then
                        Exit Do
                    End If
                    
                    If (DateDiff("s", auf2!datum, Now) < 60) Then
                        AuftragAdd% = True
                        Exit Do
                    End If
                    
                    auf2.MoveNext
                Loop
            End If
        ElseIf AbholerAuftragsNr$ > "" Then
            On Error Resume Next
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerAdoRec!KuNr)
            SQLStr$ = SQLStr$ + " AND Auftragsnr = '" + AbholerAuftragsNr$ + "'"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If Err.Number = 0 Then
              If (auf2.EOF = False) Then
                  auf2.MoveFirst
                  AuftragAdd% = True
              End If
            End If
            On Error GoTo DefErr
        End If
            
        If (AuftragAdd%) Then
            auf2.Edit
                
            If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then    '
                auf2!RezAnz = 1
                
                If (HatLieferSchein%) Then
        '            If (LieferscheinNr& > 0) Then
                    If Not (PreisSchonBerechnet) Then
                        auf2!Summe = auf2!Summe + dBesAconto#
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                        Else
    '                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
    '                        End If
    '                    End If
                    End If
        '            End If
                Else
                    auf2!Summe = auf2!Summe + dBesOffen#
                End If
            Else
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
        '            If (LieferscheinNr& > 0) Then
                        auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis
        '            Else
        '            End If
                    End If
                Else
                    auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis - dBesAconto#
                End If
            End If
            
            auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
            
            Err.Clear
            auf2.Update
        Else
            auf.Index = "Unique"
            AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
            s$ = Space(15)
            i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
            If (i% > 0) Then
                s$ = Left(s$, i%)
                If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
                    AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
                End If
            End If
                
            Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
            Do
                i% = 0
                Call clsOpTool.EanPruef(AufNr$)
                auf.Seek "=", AufNr$
                If Not (auf.NoMatch) Then i% = 1
                arch.Index = "Unique"
                arch.Seek "=", AufNr$
                If Not (arch.NoMatch) Then i% = 1
                If (i% = 1) Then
                    AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
                End If
            Loop While (i% = 1)
            i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
            
            Zeitraum$ = ""
            zahlart% = 1
            If (kAdoRec!SammelQuittung > 0) Then
                zahlart = 2
            End If
            
            If (clsOpTool.CheckNullLong(kAdoRec!LieferadresseKz) <> 0) Then
                SQLStr$ = "SELECT * FROM AdressenZusatz WHERE KundenNr =" + Str$(AbholerNummerAdoRec!KuNr)
                SQLStr$ = SQLStr$ + " AND Typ=0"
                kAdrAdoRec.Open SQLStr, KundenDB1.ActiveConn
                If (kAdrAdoRec.EOF = False) Then
                    ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Titel)))
                    ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Name)))
                    ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Station)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!HausNr)))
                    ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdrAdoRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdrAdoRec!Ort)))
                End If
                kAdrAdoRec.Close
            End If
            If (ZustName$(0) = "") And (ZustName$(1) = "") Then
                ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Titel)))
                ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Name)))
            End If
            If (ZustName$(2) = "") And (ZustName$(3) = "") Then
                ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!HausNr)))
                ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Ort)))
            End If
            If (IstMischKunde% = 0) Then
                ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!MobilPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!FonPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!FonDienst))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kAdoRec!MobilDienst))
            End If
            
            auf.AddNew
            auf!AuftragsNr = AufNr$
            auf!ProfilNr = AbholerNummerAdoRec!ProfilNr
            auf!knr = AbholerNummerAdoRec!KuNr
            auf!datum = Now
            
            LSnr$ = ""
            If (HatLieferSchein%) Then
              If Not (PreisSchonBerechnet) Then
        '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
                LSnr$ = Trim(AbholerNummerAdoRec!LieferScheinNr)
              End If
            End If
            auf!LSnr = LSnr$
                
            auf!Zahlungsart = zahlart%
            auf!Summe = 0
            If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then
                auf!RezAnz = 1
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
                        auf!Summe = dBesAconto#
        '            If (LieferscheinNr& > 0) Then
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                            auf!Summe = 0
    '                        Else
    '                            auf!Summe = tRec2!ZuzaNeu / 100
    '                        End If
    '                    Else
    '                        auf!Summe = 0
    '                    End If
        '            Else
        '                auf!Summe = 0
        '            End If
                    End If
                Else
                    auf!Summe = dBesOffen#
                End If
            Else
                auf!RezAnz = 0
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
    '                    auf!Summe = AbholerDetailRec!Preis
                        auf!Summe = clsOpTool.FNX(AbholerDetailAdoRec!lm * AbholerDetailAdoRec!Preis)
                    End If
                Else
    '                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
                    auf!Summe = clsOpTool.FNX(AbholerDetailAdoRec!lm * AbholerDetailAdoRec!Preis) - dBesAconto#
                End If
            End If
            
            prof.Index = "Unique"
            prof.Seek "=", auf!ProfilNr
            If (prof.NoMatch) Then
                auf!RezBringen = prof!RezBringen
            Else
                auf!RezBringen = 0
            End If
            
    '        auf!Positionen = 1  'ZustellAnz
            auf!Positionen = AbholerDetailAdoRec!lm  'ZustellAnz
            auf!Kundenichtda = False
            auf!kZu1 = Left$(ZustName(0), 30)
            auf!kZu2 = Left$(ZustName(1), 30)
            auf!kZu3 = Left$(ZustName(2), 30)
            auf!kZu4 = Left$(ZustName(3), 30)
            auf!kZu5 = Left$(ZustName(4), 30)
            auf!Zeitraum = Zeitraum$
            
            Err.Clear
            auf.Update
            
        End If
        
        kAdoRec.Close
        KundenDB1.CloseDB
    Else
        On Error Resume Next
        Err.Clear
        Set kDB = OpenDatabase("kunden.mdb", False, True)
        If (Err.Number <> 0) Then
            On Error GoTo DefErr
            AufDB.Close
            ProfDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        On Error GoTo DefErr
        
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(AbholerNummerAdoRec!KuNr)
        Set kRec = kDB.OpenRecordset(SQLStr$)
        If (kRec.EOF) Then
            AufDB.Close
            ProfDB.Close
            kDB.Close
            Call clsError.DefErrPop: Exit Function
        End If
        
        IstMischKunde% = kRec!MischKunde
        
        KundeGebFrei% = 0
        If Not IsNull(kRec!BefreiungBis) And Not IsEmpty(kRec!BefreiungBis) Then
            If (kRec!BefreiungBis >= Now) Then
                KundeGebFrei% = True
            End If
        End If
        
        AuftragAdd% = 0
        If (IstMischKunde% = 0) Then
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerAdoRec!KuNr)
            SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
            SQLStr$ = SQLStr$ + " AND FERTIG=0"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If (auf2.EOF = False) Then
                auf2.MoveFirst
                Do
                    If (auf2.EOF) Then
                        Exit Do
                    End If
                    
                    If (DateDiff("s", auf2!datum, Now) < 60) Then
                        AuftragAdd% = True
                        Exit Do
                    End If
                    
                    auf2.MoveNext
                Loop
            End If
        ElseIf AbholerAuftragsNr$ > "" Then
            On Error Resume Next
            SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerAdoRec!KuNr)
            SQLStr$ = SQLStr$ + " AND Auftragsnr = '" + AbholerAuftragsNr$ + "'"
            Set auf2 = AufDB.OpenRecordset(SQLStr$)
            If Err.Number = 0 Then
              If (auf2.EOF = False) Then
                  auf2.MoveFirst
                  AuftragAdd% = True
              End If
            End If
            On Error GoTo DefErr
        End If
            
        If (AuftragAdd%) Then
            auf2.Edit
                
            If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then    '
                auf2!RezAnz = 1
                
                If (HatLieferSchein%) Then
        '            If (LieferscheinNr& > 0) Then
                    If Not (PreisSchonBerechnet) Then
                        auf2!Summe = auf2!Summe + dBesAconto#
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                        Else
    '                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
    '                        End If
    '                    End If
                    End If
        '            End If
                Else
                    auf2!Summe = auf2!Summe + dBesOffen#
                End If
            Else
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
        '            If (LieferscheinNr& > 0) Then
                        auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis
        '            Else
        '            End If
                    End If
                Else
                    auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis - dBesAconto#
                End If
            End If
            
            auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
            
            Err.Clear
            auf2.Update
        Else
            auf.Index = "Unique"
            AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
            s$ = Space(15)
            i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
            If (i% > 0) Then
                s$ = Left(s$, i%)
                If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
                    AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
                End If
            End If
                
            Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
            Do
                i% = 0
                Call clsOpTool.EanPruef(AufNr$)
                auf.Seek "=", AufNr$
                If Not (auf.NoMatch) Then i% = 1
                arch.Index = "Unique"
                arch.Seek "=", AufNr$
                If Not (arch.NoMatch) Then i% = 1
                If (i% = 1) Then
                    AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
                End If
            Loop While (i% = 1)
            i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
            
            Zeitraum$ = ""
            zahlart% = 1
            If (kRec!SammelQuittung > 0) Then
                zahlart = 2
            End If
            
            If (clsOpTool.CheckNullLong(kRec!LieferadresseKz) <> 0) Then
                SQLStr$ = "SELECT * FROM AdressenZusatz WHERE KundenNr =" + Str$(AbholerNummerAdoRec!KuNr)
                SQLStr$ = SQLStr$ + " AND Typ=0"
                Set kAdrRec = kDB.OpenRecordset(SQLStr$)
                If (kAdrRec.EOF = False) Then
                    ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Titel)))
                    ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Name)))
                    ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Station)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!HausNr)))
                    ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Ort)))
                End If
            End If
            If (ZustName$(0) = "") And (ZustName$(1) = "") Then
                ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Titel)))
                ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Name)))
            End If
            If (ZustName$(2) = "") And (ZustName$(3) = "") Then
                ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kRec!HausNr)))
                ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kRec!PLZ)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Ort)))
            End If
            If (IstMischKunde% = 0) Then
                ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonPrivat))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonDienst))
                If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilDienst))
            End If
            
            auf.AddNew
            auf!AuftragsNr = AufNr$
            auf!ProfilNr = AbholerNummerAdoRec!ProfilNr
            auf!knr = AbholerNummerAdoRec!KuNr
            auf!datum = Now
            
            LSnr$ = ""
            If (HatLieferSchein%) Then
              If Not (PreisSchonBerechnet) Then
        '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
                LSnr$ = Trim(AbholerNummerAdoRec!LieferScheinNr)
              End If
            End If
            auf!LSnr = LSnr$
                
            auf!Zahlungsart = zahlart%
            auf!Summe = 0
            If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then
                auf!RezAnz = 1
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
                        auf!Summe = dBesAconto#
        '            If (LieferscheinNr& > 0) Then
    '                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
    '                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
    '                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    '                        If (tRec2.EOF) Then
    '                            auf!Summe = 0
    '                        Else
    '                            auf!Summe = tRec2!ZuzaNeu / 100
    '                        End If
    '                    Else
    '                        auf!Summe = 0
    '                    End If
        '            Else
        '                auf!Summe = 0
        '            End If
                    End If
                Else
                    auf!Summe = dBesOffen#
                End If
            Else
                auf!RezAnz = 0
                If (HatLieferSchein%) Then
                    If Not (PreisSchonBerechnet) Then
    '                    auf!Summe = AbholerDetailRec!Preis
                        auf!Summe = clsOpTool.FNX(AbholerDetailAdoRec!lm * AbholerDetailAdoRec!Preis)
                    End If
                Else
    '                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
                    auf!Summe = clsOpTool.FNX(AbholerDetailAdoRec!lm * AbholerDetailAdoRec!Preis) - dBesAconto#
                End If
            End If
            
            prof.Index = "Unique"
            prof.Seek "=", auf!ProfilNr
            If (prof.NoMatch) Then
                auf!RezBringen = prof!RezBringen
            Else
                auf!RezBringen = 0
            End If
            
    '        auf!Positionen = 1  'ZustellAnz
            auf!Positionen = AbholerDetailAdoRec!lm  'ZustellAnz
            auf!Kundenichtda = False
            auf!kZu1 = Left$(ZustName(0), 30)
            auf!kZu2 = Left$(ZustName(1), 30)
            auf!kZu3 = Left$(ZustName(2), 30)
            auf!kZu4 = Left$(ZustName(3), 30)
            auf!kZu5 = Left$(ZustName(4), 30)
            auf!Zeitraum = Zeitraum$
            
            Err.Clear
            auf.Update
            
            kDB.Close
        End If
    End If
    
'Else
'    On Error GoTo DefErr
'
'    KZUSTELL% = clsDat.FileOpen("kzustell.dat", "R", "R", 256)
'    If (KZUSTELL% = 0) Then
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    FabsErrf% = clsDat.IndexOpen("kzustell.ix1", KZUST_IDX%)
'    If (FabsErrf%) Then
'        Close #KZUSTELL%
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    MachAuftragAdo% = True
'
'    If (prof!AutoZustellung = 0) Then
'        Close #KZUSTELL%
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    clsKundenZus1.OpenDatei
'    clsKundenZus1.GetRecord (AbholerNummerAdoRec!KuNr + 1)
'    IstMischKunde% = clsKundenZus1.MischKunde
'    clsKundenZus1.CloseDatei
'    Set clsKundenZus1 = Nothing
'
'    AuftragAdd% = 0
'    If (IstMischKunde% = 0) Then
'        SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerAdoRec!KuNr)
'        SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
'        SQLStr$ = SQLStr$ + " AND FERTIG=0"
'        Set auf2 = AufDB.OpenRecordset(SQLStr$)
'        If (auf2.EOF = False) Then
'            auf2.MoveFirst
'            Do
'                If (auf2.EOF) Then
'                    Exit Do
'                End If
'
'                If (DateDiff("s", auf2!datum, Now) < 60) Then
'                    AuftragAdd% = True
'                    Exit Do
'                End If
'
'                auf2.MoveNext
'            Loop
'        End If
'    End If
'
'    If (AuftragAdd%) Then
'        Close #KZUSTELL%
'        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
'
'        auf2.Edit
'
'        If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then    '
'            auf2!RezAnz = 1
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf2!Summe = auf2!Summe + dBesAconto#
'    '            If (LieferscheinNr& > 0) Then
''                    If (Val(BesPzn$) > 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                        Else
''                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
''                        End If
''                    End If
'    '            End If
'                End If
'            Else
'                auf2!Summe = auf2!Summe + dBesOffen#
'            End If
'        Else
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis
'    '            Else
'    '            End If
'                End If
'            Else
'                auf2!Summe = auf2!Summe + AbholerDetailAdoRec!Preis - dBesAconto#
'            End If
'        End If
'
'        auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
'
'        Err.Clear
'        auf2.Update
'    Else
'        auf.Index = "Unique"
'        AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
'        s$ = Space(15)
'        i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
'        If (i% > 0) Then
'            s$ = Left(s$, i%)
'            If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
'                AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
'            End If
'        End If
'
'        Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
'        Do
'            i% = 0
'            Call clsOpTool.EanPruef(AufNr$)
'            auf.Seek "=", AufNr$
'            If Not (auf.NoMatch) Then i% = 1
'            arch.Index = "Unique"
'            arch.Seek "=", AufNr$
'            If Not (arch.NoMatch) Then i% = 1
'            If (i% = 1) Then
'                AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
'            End If
'        Loop While (i% = 1)
'        i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
'
'        zahlart% = 1
'        FabsErrf% = clsDat.IndexSearch%(KZUST_IDX%, Format(AbholerNummerAdoRec!KuNr, "00000"), recno&)
'        If (recno& > 0) Then
'            Get #KZUSTELL%, recno& + 1, kzu
'
'            zahlart% = Val(kzu.Zahlungsart)
'            ZustName$(0) = Trim(kzu.Name)
'            ZustName$(1) = Trim(kzu.zHdn)
'            ZustName$(2) = Trim(kzu.Strasse)
'            ZustName$(3) = Trim(kzu.PlzOrt)
'            ZustName$(4) = Trim(kzu.telefon)
'            Zeitraum$ = Trim(kzu.Zeitraum)
'        End If
'
'        Close #KZUSTELL%
'        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
'
'
'        clsKunden1.OpenDatei
'        clsKunden1.GetRecord (AbholerNummerAdoRec!KuNr + 1)
'
'        If (ZustName$(0) = "") And (ZustName$(1) = "") Then
'            ZustName$(0) = RTrim$(clsKunden1.Wert("name", 0))
'            ZustName$(1) = RTrim$(clsKunden1.Wert("name", 1))
'        End If
'        If (ZustName$(2) = "") And (ZustName$(3) = "") Then
'            ZustName$(2) = RTrim$(clsKunden1.Wert("name", 2))
'            ZustName$(3) = RTrim$(clsKunden1.Wert("name", 3))
'        End If
'        'If (zuName(4) = "") Then
'        '    GetDataRecord dbKZus, AktKundensatz
'        '    zuName(4) = Trim(kzus.Tel)
'        'End If
'        For i% = 0 To 3
'            Call OemToChar(ZustName$(i%), ZustName$(i%))
'        Next i%
'        Call OemToChar(Zeitraum$, Zeitraum$)
'
'        auf.AddNew
'        auf!AuftragsNr = AufNr$
'        auf!ProfilNr = AbholerNummerAdoRec!ProfilNr
'        auf!knr = AbholerNummerAdoRec!KuNr
'        auf!datum = Now
'
'        LSnr$ = ""
'        If (HatLieferSchein%) Then
'    '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
'          If Not (PreisSchonBerechnet) Then
'            LSnr$ = Trim(AbholerNummerAdoRec!LieferScheinNr)
'          End If
'        End If
'        auf!LSnr = LSnr$
'
'        auf!Zahlungsart = zahlart%
'        auf!Summe = 0
'        If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then
'            auf!RezAnz = 1
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf!Summe = dBesAconto#
'    '            If (LieferscheinNr& > 0) Then
''                    If (Val(BesPzn$) > 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                            auf!Summe = 0
''                        Else
''                            auf!Summe = tRec2!ZuzaNeu / 100
''                        End If
''                    Else
''                        auf!Summe = 0
''                    End If
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = dBesOffen#
'            End If
'        Else
'            auf!RezAnz = 0
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf!Summe = AbholerDetailAdoRec!Preis
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = AbholerDetailAdoRec!Preis - dBesAconto#
'            End If
'        End If
'
'        prof.Index = "Unique"
'        prof.Seek "=", auf!ProfilNr
'        If (prof.NoMatch) Then
'            auf!RezBringen = prof!RezBringen
'        Else
'            auf!RezBringen = 0
'        End If
'
'        auf!Positionen = 1  'ZustellAnz
'        auf!Kundenichtda = False
'        auf!kZu1 = ZustName(0)
'        auf!kZu2 = ZustName(1)
'        auf!kZu3 = ZustName(2)
'        auf!kZu4 = ZustName(3)
'        auf!kZu5 = ZustName(4)
'        auf!Zeitraum = Zeitraum$
'
'        Err.Clear
'        auf.Update
'
'        clsKunden1.CloseDatei
'        Set clsKunden1 = Nothing
'    End If
End If
    
AufDB.Close
ProfDB.Close


s$ = Space(10)
i% = GetPrivateProfileString("LIFMANAG", "Aktualisieren", "", s$, 255, "winop.ini")
If (i% > 0) Then
    i% = Val(Left(s$, i%))
Else
    i% = 0
End If
If (i% >= 32000) Then
    i% = 0
End If
i% = i% + 1
s$ = Format(i%, "0")
j% = WritePrivateProfileString("LIFMANAG", "Aktualisieren", s$, "winop.ini")

Call clsError.DefErrPop
End Function

'Function MachAuftrag%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("MachAuftrag%")
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
'Dim clsKunden1 As New clsKunden
'Dim clsKundenZus1 As New clsKundZusatz
'Dim AufDB As Database
'Dim ProfDB As Database
'Dim kDB As Database
'Dim auf As Recordset
'Dim auf2 As Recordset
'Dim prof As Recordset
'Dim arch As Recordset
'Dim kRec As Recordset
'Dim kAdrRec As Recordset
'Dim i%, j%, KZUSTELL%, KZUST_IDX%, zahlart%, ok%, IstMischKunde%, KundeGebFrei%, AuftragAdd%, HatLieferSchein%, PreisSchonBerechnet As Boolean
'Dim dBesAconto#, dBesPreis#, dBesOffen#
'Dim recno&, LieferScheinNr&
'Dim s$, AufNr$, ZustName$(4), Zeitraum$, SQLStr$, LSnr$, BesPzn$
'Dim tRec2 As Recordset
'
'
'MachAuftrag% = 0
'
'On Error Resume Next
'Err.Clear
'Set ProfDB = OpenDatabase("PROFILE.MDB")
'Set prof = ProfDB.OpenRecordset("Profile", dbOpenTable)
'If Err.Number <> 0 Then
'    On Error GoTo DefErr
'    Call clsError.DefErrPop: Exit Function
'End If
'
'prof.index = "Unique"
'prof.Seek "=", AbholerNummerRec!ProfilNr
'ok% = 0
'If (prof.NoMatch = False) Then
'    If (prof!AbhFertig) Then
'        ok% = True
'    End If
'End If
'If (ok% = 0) Then
'    On Error GoTo DefErr
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
''dBesAconto# = clsOpTool.FNX(clsOpTool.xVal(BesAconto$))
''dBesPreis# = clsOpTool.FNX(BesMenge% * clsOpTool.xVal(aw.Preis))
'dBesAconto# = AbholerDetailRec!Aconto
'dBesOffen# = AbholerDetailRec!Offen
'dBesPreis# = clsOpTool.FNX(AbholerDetailRec!lm * AbholerDetailRec!Preis)
'BesPzn$ = AbholerDetailRec!pzn
'
''LieferscheinNr& = 0
's$ = ""
'On Error Resume Next
's$ = Trim(AbholerNummerRec!LieferScheinNr)
'HatLieferSchein% = (s$ <> "")
'If HatLieferSchein% <> 0 Then
'  PreisSchonBerechnet = True
'  i% = 0
'  i% = AbholerNummerRec!AuftragsLS
'  If i% <> 0 Then PreisSchonBerechnet = False
'End If
'On Error GoTo DefErr
''If (HatLieferschein%) Then
''    LieferscheinNr& = Val(s$)
''End If
'
'ok% = True
'If (Val(AbholerDetailRec!RezeptNr) = 0) Then    '
'    'Preisunterschied kann dem Kunden nur verrechnet werden, wenn kein Lieferschein und er jetzt einen Auftrag bekommt
'    If (dBesPreis# <> dBesAconto#) And ((prof!AutoZustellung = 0) Or (HatLieferSchein%)) Then
'        ok% = 0
'    End If
'Else
'    'Wenn RezGebühr noch nicht bezahlt ist und bei der Bestellung kein Lieferschien erstellt wurde,
'    'kann sie (die RezGebühr)nur im jetzt entstehenden Auftrag verrechnet werden
'    If (dBesOffen# > 0#) And (prof!AutoZustellung = 0) And (HatLieferSchein% = 0) Then
'        ok% = 0
'    End If
'End If
'
'If (ok% = 0) Then
'    Call clsDialog.MessageBox("Abholer " + Format(AbholerNummerRec!AbholerNummer, "0") + " muss an der Kasse fertiggemacht werden!", vbInformation)
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'Err.Clear
'Set AufDB = OpenDatabase("AUFTRAG.MDB")
'Set auf = AufDB.OpenRecordset("Aufträge", dbOpenTable)
'If Err.Number <> 0 Then
'    On Error GoTo DefErr
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'If (Dir("kunbezug.mdb") <> "") Then
'    MachAuftrag% = True
'
'    If (prof!AutoZustellung = 0) Then
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    On Error Resume Next
'    Err.Clear
'    Set kDB = OpenDatabase("kunden.mdb", False, True)
'    If (Err.Number <> 0) Then
'        On Error GoTo DefErr
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'    On Error GoTo DefErr
'
'    SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(AbholerNummerRec!KuNr)
'    Set kRec = kDB.OpenRecordset(SQLStr$)
'    If (kRec.EOF) Then
'        AufDB.Close
'        ProfDB.Close
'        kDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    IstMischKunde% = kRec!MischKunde
'
'    KundeGebFrei% = 0
'    If Not IsNull(kRec!BefreiungBis) And Not IsEmpty(kRec!BefreiungBis) Then
'        If (kRec!BefreiungBis >= Now) Then
'            KundeGebFrei% = True
'        End If
'    End If
'
'    AuftragAdd% = 0
'    If (IstMischKunde% = 0) Then
'        SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
'        SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
'        SQLStr$ = SQLStr$ + " AND FERTIG=0"
'        Set auf2 = AufDB.OpenRecordset(SQLStr$)
'        If (auf2.EOF = False) Then
'            auf2.MoveFirst
'            Do
'                If (auf2.EOF) Then
'                    Exit Do
'                End If
'
'                If (DateDiff("s", auf2!datum, Now) < 60) Then
'                    AuftragAdd% = True
'                    Exit Do
'                End If
'
'                auf2.MoveNext
'            Loop
'        End If
'    End If
'
'    If (AuftragAdd%) Then
'        auf2.Edit
'
'        If (Val(AbholerDetailRec!RezeptNr) > 0) Then    '
'            auf2!RezAnz = 1
'
'            If (HatLieferSchein%) Then
'    '            If (LieferscheinNr& > 0) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf2!Summe = auf2!Summe + dBesAconto#
''                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                        Else
''                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
''                        End If
''                    End If
'                End If
'    '            End If
'            Else
'                auf2!Summe = auf2!Summe + dBesOffen#
'            End If
'        Else
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf2!Summe = auf2!Summe + AbholerDetailRec!Preis
'    '            Else
'    '            End If
'                End If
'            Else
'                auf2!Summe = auf2!Summe + AbholerDetailRec!Preis - dBesAconto#
'            End If
'        End If
'
'        auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
'
'        Err.Clear
'        auf2.Update
'    Else
'        auf.index = "Unique"
'        AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
'        s$ = Space(15)
'        i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
'        If (i% > 0) Then
'            s$ = Left(s$, i%)
'            If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
'                AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
'            End If
'        End If
'
'        Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
'        Do
'            i% = 0
'            Call clsOpTool.EanPruef(AufNr$)
'            auf.Seek "=", AufNr$
'            If Not (auf.NoMatch) Then i% = 1
'            arch.index = "Unique"
'            arch.Seek "=", AufNr$
'            If Not (arch.NoMatch) Then i% = 1
'            If (i% = 1) Then
'                AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
'            End If
'        Loop While (i% = 1)
'        i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
'
'        Zeitraum$ = ""
'        zahlart% = 1
'        If (kRec!SammelQuittung > 0) Then
'            zahlart = 2
'        End If
'
'        If (clsOpTool.CheckNullLong(kRec!LieferadresseKz) <> 0) Then
'            SQLStr$ = "SELECT * FROM AdressenZusatz WHERE KundenNr =" + Str$(AbholerNummerRec!KuNr)
'            SQLStr$ = SQLStr$ + " AND Typ=0"
'            Set kAdrRec = kDB.OpenRecordset(SQLStr$)
'            If (kAdrRec.EOF = False) Then
'                ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Titel)))
'                ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Vorname)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Name)))
'                ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Station)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!HausNr)))
'                ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kAdrRec!Plz)) + " " + Trim(clsOpTool.CheckNullStr(kAdrRec!Ort)))
'            End If
'        End If
'        If (ZustName$(0) = "") And (ZustName$(1) = "") Then
'            ZustName(0) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Anrede)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Titel)))
'            ZustName(1) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Vorname)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Name)))
'        End If
'        If (ZustName$(2) = "") And (ZustName$(3) = "") Then
'            ZustName(2) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Strasse)) + " " + Trim(clsOpTool.CheckNullStr(kRec!HausNr)))
'            ZustName(3) = Trim(Trim(clsOpTool.CheckNullStr(kRec!Plz)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Ort)))
'        End If
'        If (IstMischKunde% = 0) Then
'            ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilPrivat))
'            If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonPrivat))
'            If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!FonDienst))
'            If (ZustName(4) = "") Then ZustName(4) = Trim(clsOpTool.CheckNullStr(kRec!MobilDienst))
'        End If
'
'        auf.AddNew
'        auf!AuftragsNr = AufNr$
'        auf!ProfilNr = AbholerNummerRec!ProfilNr
'        auf!knr = AbholerNummerRec!KuNr
'        auf!datum = Now
'
'        LSnr$ = ""
'        If (HatLieferSchein%) Then
'          If Not (PreisSchonBerechnet) Then
'    '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
'            LSnr$ = Trim(AbholerNummerRec!LieferScheinNr)
'          End If
'        End If
'        auf!LSnr = LSnr$
'
'        auf!Zahlungsart = zahlart%
'        auf!Summe = 0
'        If (Val(AbholerDetailRec!RezeptNr) > 0) Then
'            auf!RezAnz = 1
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf!Summe = dBesAconto#
'    '            If (LieferscheinNr& > 0) Then
''                    If (Val(BesPzn$) > 0) And (KundeGebFrei% = 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                            auf!Summe = 0
''                        Else
''                            auf!Summe = tRec2!ZuzaNeu / 100
''                        End If
''                    Else
''                        auf!Summe = 0
''                    End If
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = dBesOffen#
'            End If
'        Else
'            auf!RezAnz = 0
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf!Summe = AbholerDetailRec!Preis
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
'            End If
'        End If
'
'        prof.index = "Unique"
'        prof.Seek "=", auf!ProfilNr
'        If (prof.NoMatch) Then
'            auf!RezBringen = prof!RezBringen
'        Else
'            auf!RezBringen = 0
'        End If
'
'        auf!Positionen = 1  'ZustellAnz
'        auf!Kundenichtda = False
'        auf!kZu1 = Left$(ZustName(0), 30)
'        auf!kZu2 = Left$(ZustName(1), 30)
'        auf!kZu3 = Left$(ZustName(2), 30)
'        auf!kZu4 = Left$(ZustName(3), 30)
'        auf!kZu5 = Left$(ZustName(4), 30)
'        auf!Zeitraum = Zeitraum$
'
'        Err.Clear
'        auf.Update
'
'        kDB.Close
'    End If
'Else
'    On Error GoTo DefErr
'
'    KZUSTELL% = clsDat.FileOpen("kzustell.dat", "R", "R", 256)
'    If (KZUSTELL% = 0) Then
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    FabsErrf% = clsDat.IndexOpen("kzustell.ix1", KZUST_IDX%)
'    If (FabsErrf%) Then
'        Close #KZUSTELL%
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    MachAuftrag% = True
'
'    If (prof!AutoZustellung = 0) Then
'        Close #KZUSTELL%
'        AufDB.Close
'        ProfDB.Close
'        Call clsError.DefErrPop: Exit Function
'    End If
'
'    clsKundenZus1.OpenDatei
'    clsKundenZus1.GetRecord (AbholerNummerRec!KuNr + 1)
'    IstMischKunde% = clsKundenZus1.MischKunde
'    clsKundenZus1.CloseDatei
'    Set clsKundenZus1 = Nothing
'
'    AuftragAdd% = 0
'    If (IstMischKunde% = 0) Then
'        SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(AbholerNummerRec!KuNr)
'        SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
'        SQLStr$ = SQLStr$ + " AND FERTIG=0"
'        Set auf2 = AufDB.OpenRecordset(SQLStr$)
'        If (auf2.EOF = False) Then
'            auf2.MoveFirst
'            Do
'                If (auf2.EOF) Then
'                    Exit Do
'                End If
'
'                If (DateDiff("s", auf2!datum, Now) < 60) Then
'                    AuftragAdd% = True
'                    Exit Do
'                End If
'
'                auf2.MoveNext
'            Loop
'        End If
'    End If
'
'    If (AuftragAdd%) Then
'        Close #KZUSTELL%
'        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
'
'        auf2.Edit
'
'        If (Val(AbholerDetailRec!RezeptNr) > 0) Then    '
'            auf2!RezAnz = 1
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf2!Summe = auf2!Summe + dBesAconto#
'    '            If (LieferscheinNr& > 0) Then
''                    If (Val(BesPzn$) > 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                        Else
''                            auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
''                        End If
''                    End If
'    '            End If
'                End If
'            Else
'                auf2!Summe = auf2!Summe + dBesOffen#
'            End If
'        Else
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf2!Summe = auf2!Summe + AbholerDetailRec!Preis
'    '            Else
'    '            End If
'                End If
'            Else
'                auf2!Summe = auf2!Summe + AbholerDetailRec!Preis - dBesAconto#
'            End If
'        End If
'
'        auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
'
'        Err.Clear
'        auf2.Update
'    Else
'        auf.index = "Unique"
'        AufNr$ = "97" + Format(Now, "YYMMDD") + "00010"
'        s$ = Space(15)
'        i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, "dp.ini")
'        If (i% > 0) Then
'            s$ = Left(s$, i%)
'            If (Left(s$, 2) = "97") And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
'                AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
'            End If
'        End If
'
'        Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
'        Do
'            i% = 0
'            Call clsOpTool.EanPruef(AufNr$)
'            auf.Seek "=", AufNr$
'            If Not (auf.NoMatch) Then i% = 1
'            arch.index = "Unique"
'            arch.Seek "=", AufNr$
'            If Not (arch.NoMatch) Then i% = 1
'            If (i% = 1) Then
'                AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
'            End If
'        Loop While (i% = 1)
'        i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, "dp.ini")
'
'        zahlart% = 1
'        FabsErrf% = clsDat.IndexSearch%(KZUST_IDX%, Format(AbholerNummerRec!KuNr, "00000"), recno&)
'        If (recno& > 0) Then
'            Get #KZUSTELL%, recno& + 1, kzu
'
'            zahlart% = Val(kzu.Zahlungsart)
'            ZustName$(0) = Trim(kzu.Name)
'            ZustName$(1) = Trim(kzu.zHdn)
'            ZustName$(2) = Trim(kzu.Strasse)
'            ZustName$(3) = Trim(kzu.PlzOrt)
'            ZustName$(4) = Trim(kzu.telefon)
'            Zeitraum$ = Trim(kzu.Zeitraum)
'        End If
'
'        Close #KZUSTELL%
'        FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
'
'
'        clsKunden1.OpenDatei
'        clsKunden1.GetRecord (AbholerNummerRec!KuNr + 1)
'
'        If (ZustName$(0) = "") And (ZustName$(1) = "") Then
'            ZustName$(0) = RTrim$(clsKunden1.Wert("name", 0))
'            ZustName$(1) = RTrim$(clsKunden1.Wert("name", 1))
'        End If
'        If (ZustName$(2) = "") And (ZustName$(3) = "") Then
'            ZustName$(2) = RTrim$(clsKunden1.Wert("name", 2))
'            ZustName$(3) = RTrim$(clsKunden1.Wert("name", 3))
'        End If
'        'If (zuName(4) = "") Then
'        '    GetDataRecord dbKZus, AktKundensatz
'        '    zuName(4) = Trim(kzus.Tel)
'        'End If
'        For i% = 0 To 3
'            Call OemToChar(ZustName$(i%), ZustName$(i%))
'        Next i%
'        Call OemToChar(Zeitraum$, Zeitraum$)
'
'        auf.AddNew
'        auf!AuftragsNr = AufNr$
'        auf!ProfilNr = AbholerNummerRec!ProfilNr
'        auf!knr = AbholerNummerRec!KuNr
'        auf!datum = Now
'
'        LSnr$ = ""
'        If (HatLieferSchein%) Then
'    '        LSnr$ = Format(Now, "YY") + "/" + Format(LieferscheinNr&, "00000") + "V"
'          If Not (PreisSchonBerechnet) Then
'            LSnr$ = Trim(AbholerNummerRec!LieferScheinNr)
'          End If
'        End If
'        auf!LSnr = LSnr$
'
'        auf!Zahlungsart = zahlart%
'        auf!Summe = 0
'        If (Val(AbholerDetailRec!RezeptNr) > 0) Then
'            auf!RezAnz = 1
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'                    auf!Summe = dBesAconto#
'    '            If (LieferscheinNr& > 0) Then
''                    If (Val(BesPzn$) > 0) Then
''                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
''                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
''                        If (tRec2.EOF) Then
''                            auf!Summe = 0
''                        Else
''                            auf!Summe = tRec2!ZuzaNeu / 100
''                        End If
''                    Else
''                        auf!Summe = 0
''                    End If
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = dBesOffen#
'            End If
'        Else
'            auf!RezAnz = 0
'            If (HatLieferSchein%) Then
'                If Not (PreisSchonBerechnet) Then
'    '            If (LieferscheinNr& > 0) Then
'                    auf!Summe = AbholerDetailRec!Preis
'    '            Else
'    '                auf!Summe = 0
'    '            End If
'                End If
'            Else
'                auf!Summe = AbholerDetailRec!Preis - dBesAconto#
'            End If
'        End If
'
'        prof.index = "Unique"
'        prof.Seek "=", auf!ProfilNr
'        If (prof.NoMatch) Then
'            auf!RezBringen = prof!RezBringen
'        Else
'            auf!RezBringen = 0
'        End If
'
'        auf!Positionen = 1  'ZustellAnz
'        auf!Kundenichtda = False
'        auf!kZu1 = ZustName(0)
'        auf!kZu2 = ZustName(1)
'        auf!kZu3 = ZustName(2)
'        auf!kZu4 = ZustName(3)
'        auf!kZu5 = ZustName(4)
'        auf!Zeitraum = Zeitraum$
'
'        Err.Clear
'        auf.Update
'
'        clsKunden1.CloseDatei
'        Set clsKunden1 = Nothing
'    End If
'End If
'
'AufDB.Close
'ProfDB.Close
'
'
's$ = Space(10)
'i% = GetPrivateProfileString("LIFMANAG", "Aktualisieren", "", s$, 255, "winop.ini")
'If (i% > 0) Then
'    i% = Val(Left(s$, i%))
'Else
'    i% = 0
'End If
'If (i% >= 32000) Then
'    i% = 0
'End If
'i% = i% + 1
's$ = Format(i%, "0")
'j% = WritePrivateProfileString("LIFMANAG", "Aktualisieren", s$, "winop.ini")
'
'Call clsError.DefErrPop
'End Function

Sub SpeicherBezug()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherBezug")
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
Dim i%, j%, iFixiert%, iZeit%, iRezeptArt%
Dim LaufNr&, FabsRecno&
Dim fAnzPackungen!, MaxBezugRec!
Dim Preis#
Dim pzn$, key$, txt$, h$, SQLStr$
Dim KunBezugDB As Database
Dim KunBezugRec As Recordset

'If (Dir("KunBezug.mdb") = "") Then
'    Call clsError.DefErrPop: Exit Sub
'End If
'
'If (Dir("verkauf.mdb") <> "") Then
'    Call clsError.DefErrPop: Exit Sub
'End If
'
'If (AbholerNummerRec!KuNr > 0) Then
'    Set KunBezugDB = OpenDatabase("KunBezug.mdb", False, False)
'    Set KunBezugRec = KunBezugDB.OpenRecordset("Bezug", dbOpenTable)
'    KunBezugRec.Index = "Unique"
'
'    pzn$ = clsOpTool.PznString(AbholerDetailRec!pzn)
'    fAnzPackungen! = AbholerDetailRec!lm
'
''    If (Val(AbholerDetailRec!RezeptNr) <> 0) Then
''        If (Dir("verkauf.mdb") = "") Then
''            Preis# = AbholerDetailRec!Offen
''        ElseIf (AbholerDetailRec!Aconto <> 0) Then
''            Preis# = AbholerDetailRec!Aconto
''        Else
''            Preis# = AbholerDetailRec!Offen
''        End If
''        If (fAnzPackungen > 1) Then
''            Preis# = Preis# / AbholerDetailRec!bm
''        End If
''    Else
''        Preis# = AbholerDetailRec!Preis
''    End If
''    If ((Abs(fAnzPackungen! * Preis#) <> 0) Or (pzn$ <> "9999999")) Then
'
'    Preis# = AbholerDetailRec!Offen + AbholerDetailRec!Aconto
'    If (fAnzPackungen > 1) Then
'        Preis# = Preis# / AbholerDetailRec!bm
'    End If
'
'    If (fAnzPackungen <> 0) Then
'        KunBezugRec.AddNew
'
'        KunBezugRec!pzn = Val(pzn$)
'        KunBezugRec!text = Trim(AbholerDetailRec!text)
'        KunBezugRec!menge = Trim(AbholerDetailRec!menge)
'        KunBezugRec!einheit = Trim(AbholerDetailRec!einheit)
'
'        KunBezugRec!Anzahl = fAnzPackungen!
'
'        KunBezugRec!Preis = 0
'        KunBezugRec!Zuzahlung = 0
'        If (Val(AbholerDetailRec!RezeptNr) <> 0) Then
'            KunBezugRec!Zuzahlung = Abs(Preis#) * fAnzPackungen!
'        Else
'            KunBezugRec!Preis = Abs(Preis#)
'        End If
'
'        KunBezugRec!MwStCode = AbholerDetailRec!mw
'        KunBezugRec!Warengruppe = AbholerDetailRec!wg
'
'        KunBezugRec!AbholNr = AbholerDetailRec!AbholerNummer
'        KunBezugRec!AbholId = AbholerDetailRec!AbholerInd
'
'        KunBezugRec!Computer = Val(Para1.user)
'        KunBezugRec!Seite = 0
'        KunBezugRec!Bediener = 0
'        KunBezugRec!Typ = 0
'
'        KunBezugRec!MarkierungsKZ = " "
'        KunBezugRec!Beleg = " "
'
'        KunBezugRec!RezeptNr = AbholerDetailRec!RezeptNr
'
'        iRezeptArt = AbholerDetailRec!RezeptArt
'        If (iRezeptArt = 1) Then
'            iRezeptArt = 2
'        ElseIf (iRezeptArt = 2) Then
'            iRezeptArt = 1
'        End If
'        KunBezugRec!RezeptArt = iRezeptArt
'
'        KunBezugRec!KundenNr = AbholerNummerRec!KuNr
'
'        h$ = Format(Now, "DD.MM.YYYY")
'        iZeit% = Val(Format(Now, "HHMM"))
'        KunBezugRec!datum = DateValue(h$) + TimeValue(Format(iZeit \ 100, "00") + ":" + Format(iZeit Mod 100, "00"))
'
'        KunBezugRec!BezBeleg = Format(Now, "YY") + "/" + "W" + Format(Now, "HHMMSS")
'        KunBezugRec!Bezzeile = 1
'
'        KunBezugRec!Absender = 0
'        KunBezugRec!gesendet = 0
'
'    '''''''''''''''''
'        KunBezugRec!AbgleichKZ = 0
'        KunBezugRec!PhCare = 0
'        KunBezugRec!ZahlKundenNr = AbholerNummerRec!KuNr
'
'        KunBezugRec!AnzeigeKZ = True
'        KunBezugRec!EndeT = 4
'
'        KunBezugRec!Rabattfähig = 0
'        KunBezugRec!angefordert = 0
'        KunBezugRec!Hilfsmittel = 0
'
'    '            KunBezugRec!ZuzaFlag = 0
'        KunBezugRec!RezRausKZ = 0
'        KunBezugRec!Zustellung = 0
'
'        KunBezugRec!Arzt = 0
'
'        KunBezugRec!PartnerAbh = 0
'        KunBezugRec!LieferScheinNr = " "
'        KunBezugRec!Fortsetzung = False
'        KunBezugRec!FremdWährung = 0
'        KunBezugRec!FremdZahlBetrag = 0
'        KunBezugRec!PunkteArt = 0
'        KunBezugRec!PaybackPunkte = 0
'        KunBezugRec!PaybackSonder = 0
'        KunBezugRec!PaybackFremd = False
'        'v!StornoKunde
'
'        If Val(pzn) > 0 And Val(pzn) <> 9999999 Then
'            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
'            If (TaxeAdoDBok) Then
'                On Error Resume Next
'                TaxeAdoRec.Close
'                Err.Clear
'                On Error GoTo DefErr
'                TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
'                If (TaxeAdoRec.EOF = False) Then
'                    KunBezugRec!text = TaxeAdoRec!Name
'                    KunBezugRec!menge = TaxeAdoRec!menge
'                    KunBezugRec!einheit = TaxeAdoRec!einheit
'                    KunBezugRec!AbgabeBest = TaxeAdoRec!AbgabeBest
'                    KunBezugRec!LauerWG = TaxeAdoRec!Warengruppe
'                    KunBezugRec!HimiVerbrauch = (TaxeAdoRec!HimiVerbrauch > 0)
'                    If Not (IsNull(TaxeAdoRec!NGrösse)) Then
'                        KunBezugRec!NGrösse = Val(Mid(TaxeAdoRec!NGrösse, 2, 1))
'                    End If
'                    If TaxeAdoRec!BTMKz > 0 Then KunBezugRec!Betäubungsmittel = True
'                    KunBezugRec!ZuzZeilenwertKZ = (TaxeAdoRec!ZuzZeilenwertKZ > 0)
'                    KunBezugRec!ZzfreiFbKz = (TaxeAdoRec!ZzfreiFbKz > 0)
'                    KunBezugRec!EK = (TaxeAdoRec!EK / 100#)
'                    KunBezugRec!vk = (TaxeAdoRec!vk / 100#)
'                    KunBezugRec!FB = (TaxeAdoRec!FESTBETRAG / 100#)
'                End If
'            Else
'                Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                If (TaxeRec.EOF = False) Then
'                    KunBezugRec!text = TaxeRec!Name
'                    KunBezugRec!menge = TaxeRec!menge
'                    KunBezugRec!einheit = TaxeRec!einheit
'                    KunBezugRec!AbgabeBest = TaxeRec!AbgabeBest
'                    KunBezugRec!LauerWG = TaxeRec!Warengruppe
'                    KunBezugRec!HimiVerbrauch = (TaxeRec!HimiVerbrauch > 0)
'                    If Not (IsNull(TaxeRec!NGrösse)) Then
'                        KunBezugRec!NGrösse = Val(Mid(TaxeRec!NGrösse, 2, 1))
'                    End If
'                    If TaxeRec!BTMKz > 0 Then KunBezugRec!Betäubungsmittel = True
'                    KunBezugRec!ZuzZeilenwertKZ = (TaxeRec!ZuzZeilenwertKZ > 0)
'                    KunBezugRec!ZzfreiFbKz = (TaxeRec!ZzfreiFbKz > 0)
'                    KunBezugRec!EK = (TaxeRec!EK / 100#)
'                    KunBezugRec!vk = (TaxeRec!vk / 100#)
'                    KunBezugRec!FB = (TaxeRec!FESTBETRAG / 100#)
'                End If
'            End If
'
'            If (ArtikelDBok) Then
'                SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
'                FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
'            Else
'                FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
'                If (FabsErrf% = 0) Then
'                    Ast1.GetRecord (FabsRecno& + 1)
'                End If
'            End If
'
'            If (FabsErrf% = 0) Then
'                KunBezugRec!text = Ast1.kurz
'                KunBezugRec!menge = Ast1.meng
'                KunBezugRec!einheit = Ast1.meh
'
'                If (ArtikelDBok) Then
'                    FabsErrf = Abs(ArtikelRec!LagerKz = 0)
'                Else
'                    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
'                End If
'                If (FabsErrf% = 0) Then
'                    KunBezugRec!Lagerbuchen = True
'                    KunBezugRec!StatistikBuchen = True
'                End If
'            End If
'        End If
'
'        KunBezugRec.Update
'    End If
'
'    KunBezugDB.Close
'End If

Call clsError.DefErrPop
End Sub

Sub BesorgerLoeschen(Nummer%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BesorgerLoeschen")
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
Dim i%, erg%, Belegt%, KisteFertig%, KisteStatus%, BesorgerStatus%, s%, NurFertigeBesorger%
Dim h$, Arbeit$

If (Abs(Nummer%) > 999) Then
    Call clsError.DefErrPop: Exit Sub
End If
If (Nummer% < 0) Then
    Call clsError.DefErrPop: Exit Sub
End If

Set Kiste1 = New clsKiste
Kiste1.OpenDatei
Belegt% = Kiste1.Belegt(Nummer%)
If (Belegt%) Then
    Call Kiste1.GetKiste(Nummer%)
    If (Kiste1.KistenStatus = 4) Then
        Arbeit$ = Kiste1.Arbeit
        '* alles abgeholt
        '* wenn das nur Besorger waren, dann entfernen
        NurFertigeBesorger% = True
        For i% = 1 To 5
            s% = Asc(Mid$(Arbeit$, i%, 1))
            If (i% <> 4) And (s% <> 0) Then
                NurFertigeBesorger% = 0
            End If
        Next i%
        If (NurFertigeBesorger%) Then
            '* AnfBlocks ausleeren
            For i% = 0 To 9
                erg% = Kiste1.ClearInhalt(i%)
            Next i%

            Kiste1.ClearKiste (Nummer%)
        End If
    End If
End If
    
Kiste1.CloseDatei
Set Kiste1 = Nothing

Call clsError.DefErrPop
End Sub

Function TesteBesorgerBstatus$(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TesteBesorgerBstatus$")
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
Dim i%, lief%, m%, dAnzFields%, IstMerkzettelMdb%
Dim FabsRecno&
Dim h$, h2$, ret$, SQLStr$

ret$ = ""

If (Val(pzn$) = 9999999) Then
    TesteBesorgerBstatus$ = ret$
    Call clsError.DefErrPop: Exit Function
End If

FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ret$ = "Letzte Lieferung:  "
    Ass1.GetRecord (FabsRecno& + 1)
    lief% = Ass1.llief(4)
    If (LieferantenDBok) Then
        h2 = ""
        SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
        LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
        If (LieferantenRec.RecordCount <> 0) Then
            h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
        End If
        LieferantenRec.Close
        ret$ = ret$ + h2$ + " "
    ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
        Lif1.GetRecord (lief% + 1)
        h2$ = Lif1.kurz
        Call OemToChar(h2$, h2$)
        ret$ = ret$ + h2$ + " "
    End If
    ret$ = ret$ + Format(lief%, "(###) ")
    
    h2$ = "geliefert"
    ret$ = ret$ + h2$ + " am "
    
    h2$ = clsOpTool.CVDatum2$(Ass1.lld)
    ret$ = ret$ + Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2)
    ret$ = ret$ + "   " + Str$(Ass1.llm) + " Stk"
Else
    On Error Resume Next
    dAnzFields% = MerkzettelDB.TableDefs("Merkzettel").Fields.Count
    IstMerkzettelMdb% = (Err = 0)
    On Error GoTo DefErr
    
    If (IstMerkzettelMdb%) Then
        SQLStr$ = "SELECT * FROM Merkzettel WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
        Set MerkzettelRec = MerkzettelDB.OpenRecordset(SQLStr$)
        If (MerkzettelRec.RecordCount > 0) Then
            ret$ = "Merkzettel:  "
            lief% = MerkzettelRec!lief
            If (LieferantenDBok) Then
                h2 = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                End If
                LieferantenRec.Close
                ret$ = ret$ + h2$ + " "
            ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                Lif1.GetRecord (lief% + 1)
                h2$ = Lif1.kurz
                Call OemToChar(h2$, h2$)
                ret$ = ret$ + h2$ + " "
            End If
            ret$ = ret$ + Format(lief%, "(###) ")
            
            m% = MerkzettelRec!lm
            If (m% = 0) Then
                h2$ = "angefragt"
            Else
                h2$ = "besorgt"
            End If
            ret$ = ret$ + h2$ + " am "
            
            ret$ = ret$ + Format(MerkzettelRec!LieferDatum, "DD.MM.YY")
            ret$ = ret$ + "   " + Str$(m%) + " Stk"
        End If
    Else
        FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ret$ = "Merkzettel:  "
            Besorgt1.GetRecord (FabsRecno& + 1)
            lief% = Besorgt1.lief
            If (LieferantenDBok) Then
                h2 = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(lief)
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    h2$ = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                End If
                LieferantenRec.Close
                ret$ = ret$ + h2$ + " "
            ElseIf (lief% > 0) And (lief% <= Lif1.AnzRec) Then
                Lif1.GetRecord (lief% + 1)
                h2$ = Lif1.kurz
                Call OemToChar(h2$, h2$)
                ret$ = ret$ + h2$ + " "
            End If
            ret$ = ret$ + Format(lief%, "(###) ")
            
            m% = Besorgt1.lm
            If (m% = 0) Then
                h2$ = "angefragt"
            Else
                h2$ = "besorgt"
            End If
            ret$ = ret$ + h2$ + " am "
            
            h2$ = clsOpTool.CVDatum2$(Besorgt1.dt)
            ret$ = ret$ + Mid$(h2$, 7, 2) + "." + Mid$(h2$, 5, 2) + "." + Mid$(h2$, 3, 2)
            ret$ = ret$ + "   " + Str$(m%) + " Stk"
        End If
    End If
End If

TesteBesorgerBstatus$ = ret$

Call clsError.DefErrPop
End Function

Sub ZeigeAbholerEinzeln(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeAbholerEinzeln")
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
Dim i%, j%, row%, MaxInfoInd%, iToggle%
Dim ToggleFarbe&
Dim tx$, tx2$, RezTxt$, SQLStr$

If (AbholerSQL) Then
    SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerAdoRec!AbholerNummer)
    SQLStr$ = SQLStr$ + " AND AbholerInd=" + Str$(ind% + 1)
    On Error Resume Next
    AbholerDetailAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    AbholerDetailAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
    
    If (AbholerDetailAdoRec.EOF) Then
        Call clsError.DefErrPop: Exit Sub
    End If
    
    With frmAbholer.flxAbholerEinzeln
        If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then
            tx$ = "offene Gebühr"
            tx2$ = Format(AbholerDetailAdoRec!Offen, "0.00")
        Else
            tx$ = "Aconto"
            tx2$ = Format(AbholerDetailAdoRec!Aconto, "0.00")
        End If
        .TextMatrix(0, 1) = Trim(AbholerDetailAdoRec!RezeptNr)
        .TextMatrix(1, 0) = tx$
        
        .TextMatrix(1, 1) = tx2$
        .TextMatrix(2, 1) = Str$(AbholerDetailAdoRec!lm)
        
        If (Val(AbholerDetailAdoRec!RezeptNr) > 0) Then
            tx$ = ""
        Else
            tx$ = Format(AbholerDetailAdoRec!Preis * AbholerDetailAdoRec!lm, "0.00")
        End If
        .TextMatrix(3, 1) = tx$
        
        .TextMatrix(4, 1) = TesteBesorgerBstatus(clsOpTool.PznString(AbholerDetailAdoRec!pzn))
    End With
        
    SQLStr$ = "SELECT * FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailAdoRec!Id)
    SQLStr$ = SQLStr$ + " ORDER BY InfoInd"
    On Error Resume Next
    AbholerInfoAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    AbholerInfoAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
    With frmAbholer.flxAbholerInfo
        .Rows = 0
    
        Do
            If (AbholerInfoAdoRec.EOF) Then
                Exit Do
            End If
                
            MaxInfoInd% = AbholerInfoAdoRec!InfoInd
                
            .Rows = .Rows + 1
            row% = .Rows - 1
            
            .TextMatrix(row%, 0) = Format(AbholerInfoAdoRec!AngelegtAm, "DD.MM.YYYY HH:MM")
'
            tx$ = Trim$(AbholerInfoAdoRec!text)
            .TextMatrix(row%, 1) = tx$
                
            .TextMatrix(row%, 2) = Format(AbholerInfoAdoRec!InfoInd, "0")
            
            AbholerInfoAdoRec.MoveNext
        Loop
        If (.Rows < 6) Then
            .Rows = 6
        End If
    
        For i% = (.Rows - 1) To (.FixedRows + 1) Step -1
            If (.TextMatrix(i%, 0) = .TextMatrix(i% - 1, 0)) Then
                For j% = 0 To 0
                    .TextMatrix(i%, j%) = ""
                Next j%
            End If
        Next i%
        
        If (iNewLine = 0) Then
            ToggleFarbe& = RGB(190, 190, 190)
            iToggle% = 0
            For i% = .FixedRows To (.Rows - 1)
                If (.TextMatrix(i%, 0) <> "") Then
                    iToggle% = Not (iToggle%)
                End If
                
                If (iToggle%) Then
                    .FillStyle = flexFillRepeat
                    .row = i%
                    .col = 0
                    .RowSel = .row
                    .ColSel = .Cols - 1
                    .CellBackColor = ToggleFarbe&   'wpara.FarbeDunklerBereich
                    .FillStyle = flexFillSingle
                End If
            Next i%
        End If
        
        .row = 0
        .col = 1
        .ColSel = .Cols - 1
    End With
ElseIf (AbholerMdb%) Then
    SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer=" + Str$(AbholerNummerRec!AbholerNummer)
    SQLStr$ = SQLStr$ + " AND AbholerInd=" + Str$(ind% + 1)
    Set AbholerDetailRec = AbholerDB.OpenRecordset(SQLStr$)
        
    If (AbholerDetailRec.EOF) Then
        Call clsError.DefErrPop: Exit Sub
    End If
    
    With frmAbholer.flxAbholerEinzeln
        If (Val(AbholerDetailRec!RezeptNr) > 0) Then
            tx$ = "offene Gebühr"
            tx2$ = Format(AbholerDetailRec!Offen, "0.00")
        Else
            tx$ = "Aconto"
            tx2$ = Format(AbholerDetailRec!Aconto, "0.00")
        End If
        .TextMatrix(0, 1) = Trim(AbholerDetailRec!RezeptNr)
        .TextMatrix(1, 0) = tx$
        
        .TextMatrix(1, 1) = tx2$
        .TextMatrix(2, 1) = Str$(AbholerDetailRec!lm)
        
        If (Val(AbholerDetailRec!RezeptNr) > 0) Then
            tx$ = ""
        Else
            tx$ = Format(AbholerDetailRec!Preis * AbholerDetailRec!lm, "0.00")
        End If
        .TextMatrix(3, 1) = tx$
        
        .TextMatrix(4, 1) = TesteBesorgerBstatus(clsOpTool.PznString(AbholerDetailRec!pzn))
    End With
        
    SQLStr$ = "SELECT * FROM AbholerInfo WHERE AbholerDetailId=" + Str$(AbholerDetailRec!Id)
    SQLStr$ = SQLStr$ + " ORDER BY InfoInd"
    Set AbholerInfoRec = AbholerDB.OpenRecordset(SQLStr$)
    
    With frmAbholer.flxAbholerInfo
        .Rows = 0
    
        Do
            If (AbholerInfoRec.EOF) Then
                Exit Do
            End If
                
            MaxInfoInd% = AbholerInfoRec!InfoInd
                
            .Rows = .Rows + 1
            row% = .Rows - 1
            
            .TextMatrix(row%, 0) = Format(AbholerInfoRec!AngelegtAm, "DD.MM.YYYY HH:MM")
'
'            If (AbholerInfoRec!AngelegtVon > 0) Then
'                .TextMatrix(row%, 1) = Trim(para.Personal(AbholerInfoRec!AngelegtVon))
'            End If
'
'            If (AbholerInfoRec!Status >= 0) Then
'                .TextMatrix(row%, 2) = ProgrammTypStr$(AbholerInfoRec!Status)
'            End If
                
            tx$ = Trim$(AbholerInfoRec!text)
            .TextMatrix(row%, 1) = tx$
                
            .TextMatrix(row%, 2) = Format(AbholerInfoRec!InfoInd, "0")
            
            AbholerInfoRec.MoveNext
        Loop
        If (.Rows < 6) Then
            .Rows = 6
        End If
    
        For i% = (.Rows - 1) To (.FixedRows + 1) Step -1
            If (.TextMatrix(i%, 0) = .TextMatrix(i% - 1, 0)) Then
                For j% = 0 To 0
                    .TextMatrix(i%, j%) = ""
                Next j%
            End If
        Next i%
        
        If (iNewLine = 0) Then
            ToggleFarbe& = RGB(190, 190, 190)
            iToggle% = 0
            For i% = .FixedRows To (.Rows - 1)
                If (.TextMatrix(i%, 0) <> "") Then
                    iToggle% = Not (iToggle%)
                End If
                
                If (iToggle%) Then
                    .FillStyle = flexFillRepeat
                    .row = i%
                    .col = 0
                    .RowSel = .row
                    .ColSel = .Cols - 1
                    .CellBackColor = ToggleFarbe&   'wpara.FarbeDunklerBereich
                    .FillStyle = flexFillSingle
                End If
            Next i%
        End If
        
        .row = 0
        .col = 1
        .ColSel = .Cols - 1
    End With
Else
    With frmAbholer.flxAbholerEinzeln
        Kiste1.GetInhalt (ind%)
        
        tx$ = "Aconto"
        RezTxt$ = ""
        If (Kiste1.RezeptNr > 0) Then
            RezTxt$ = clsOpTool.Bcd2ascii(Kiste1.RezeptEan, 12)
            RezTxt$ = RezTxt$ + "0": Call clsOpTool.EanPruef(RezTxt$)
            tx$ = "offene Gebühr"
        End If
        .TextMatrix(0, 1) = RezTxt$
        .TextMatrix(1, 0) = tx$
        
        .TextMatrix(1, 1) = LTrim$(Kiste1.Aconto)
        .TextMatrix(2, 1) = Str$(Kiste1.menge)
        
        If (Kiste1.RezeptNr > 0) Then
            tx$ = ""
        Else
            tx$ = Format(Kiste1.Preis * Kiste1.menge, "0.00")
        End If
        .TextMatrix(3, 1) = tx$
        
        .TextMatrix(4, 1) = TesteBesorgerBstatus(Kiste1.pzn)
        
        For i% = 5 To 10
            .TextMatrix(i%, 1) = RTrim$(Kiste1.InfoText(i% - 2))
            frmAbholer.txtZusatz(i% - 5) = .TextMatrix(i%, 1)
        Next i%
    End With
End If

With frmAbholer
    tx$ = Trim$(.flxAbholerGlobal.TextMatrix(5 + ind%, 2))
    .cmdF5.Enabled = True
    If (tx$ = "ABGEHOLT") Then
        .cmdF5.Caption = "Fertig (F5)"
    Else
        .cmdF5.Caption = "Abgeholt (F5)"
    End If
    If (iNewLine) Then
        .nlcmdF5.Enabled = True
        If (tx$ = "ABGEHOLT") Then
            .nlcmdF5.Caption = "Fertig"
        Else
            .nlcmdF5.Caption = "Abgeholt"
        End If
    End If
End With

Call clsError.DefErrPop
End Sub

Function AbholerNummer%(Nummer%, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AbholerNummer%")
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
'Dim i%, j%, satz1%, satz%, header%, bit%, BitMaske%, frei%, KisteFertig%, allesunfertig%, menge%
'Dim text%, InfoText%, PznGefunden%, Belegt%, BlockNummer%
'Dim KistenSatz&
'Dim Preis#
'Dim ch$, Arttext$, aiPzn$, Aconto$, txt$, wg$, mText$, mx$, h$, h1$, h2$
'
'AbholerNummer% = False
'
'If Abs(Nummer%) > 999 Then
'    Call clsError.DefErrPop: Exit Function
'End If
'If Nummer% < 0 Then
'    Call clsError.DefErrPop: Exit Function
'End If
'
'Kiste1.OpenDatei
'Belegt% = Kiste1.Belegt(Nummer%)
'If (Belegt%) Then
'    BlockNummer% = Kiste1.PznInKiste(Nummer%, pzn$)
'    If (BlockNummer% >= 0) Then
'        Kiste1.GetInhalt (BlockNummer%)
'        AbholerNummer% = True
'    End If
'End If
'
'Kiste1.CloseDatei
'
Call clsError.DefErrPop
End Function

Sub SucheBesorgerInVk(AbholNr%, PersCode%, KundenNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SucheBesorgerInVk")
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
Dim i%, j%, VERKAUF%
Dim von&, bis&, Index&, VerkaufMax&, ret&
Dim h$, h2$, s$, SuchDatum$, SuchZeit$, Such$, SuchKiste$

PersCode% = 0
KundenNr = 0

SuchDatum$ = Left$(Kiste1.VonWann, 2)
SuchZeit$ = clsOpTool.MKI(Asc(Mid$(Kiste1.VonWann, 3, 1)) * 100 + Asc(Mid$(Kiste1.VonWann, 4, 1)))
Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)

SuchKiste$ = Format(AbholNr%, "0")

'For i% = 1 To 4
'    Debug.Print Asc(Mid$(Such$, i%, 1));
'Next i%
'Debug.Print
'Debug.Print
                
h$ = "verkauf.dat"
VERKAUF% = clsDat.FileOpen(h$, "R", "R", 128)

VerkaufMax& = (LOF(VERKAUF%) / 128) - 1

For j% = 0 To 1
    If (j% = 0) Then
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)
    Else
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Chr$(Asc(Mid$(SuchZeit$, 1, 1)) - 1)
    End If
    
    h$ = Such$
    GoSub KonvVerkStr
    Such$ = h$
    
    von& = 1&
    bis& = VerkaufMax&
    
    Do While (von& <= bis&)
        Index& = (von& + bis&) \ 2
        Get #VERKAUF%, Index& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)

        h$ = s$
        GoSub KonvVerkStr
        s$ = h$
    
        If (Such$ = s$) Then
            ret& = Index&
            Exit Do
        ElseIf (Such$ < s$) Then
            bis& = Index& - 1
        Else
            von& = Index& + 1
        End If
    
    Loop
    
    If (ret& >= 1) Then Exit For
Next j%

If (ret& >= 1) Then
    Do
        ret& = ret& - 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        
        h$ = s$
        GoSub KonvVerkStr
        s$ = h$
        
        If (Such$ <> s$) Then
            Exit Do
        End If
    Loop
    
    Do
        ret& = ret& + 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        
        h$ = s$
        GoSub KonvVerkStr
        s$ = h$
    
        If (Such$ <> s$) Then
            Exit Do
        End If
        If (InStr(VkRec.text, "Abhol-Nr") > 0) And (InStr(VkRec.text, SuchKiste$) > 0) Then
            KundenNr = VkRec.knr
            PersCode% = VkRec.PersCode
            Exit Do
        End If
    Loop
End If

Close #VERKAUF%
Call clsError.DefErrPop: Exit Sub

KonvVerkStr:
h2 = ""
For i% = 1 To Len(h$)
    h2$ = h2$ + Right$(Space$(3) + Str$(Asc(Mid$(h$, i%, 1))), 3)
Next i%
h$ = h2$
Return

Call clsError.DefErrPop
End Sub

Sub ZeigeAngebote(pzn$, txt$, angModus%, angInd%, angTemporaer As Byte, angManuell As Byte, angY%, Optional angBm% = 0, Optional angNm% = 0, Optional angLief%, Optional AngAep#, Optional angZr!, Optional angHerst$ = "", Optional angGhRuf$ = "", Optional angGhPriorität$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeAngebote")
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
Dim iAng%, bm%, iAngLen%, VglInd%, NeuSpeichern%
Dim mp&, lRecs&
Dim FabsRecno&
Dim sAng1$, sAng2$

AngebotDirektEingabe% = False
If (angInd% = -999) Then
    AngebotDirektEingabe% = True
    angInd% = -1
End If

AngebotPzn$ = pzn$
AngebotModus% = angModus%
AngebotInd% = angInd%
AngebotAuswahl% = True
AngebotBm% = angBm%
AngebotNm% = angNm%
AngebotY% = angY%
AngebotActLief% = angLief%
AngebotZr! = angZr!
AngebotTemporaer = angTemporaer
AngebotManuell = angManuell
AngebotHerst$ = angHerst$
AngebotGhRuf$ = angGhRuf$
AngebotGhPriorität$ = angGhPriorität$

IstBewertung% = False
    
Load frmAngebote
With frmAngebote
    .Caption = "Angebote - " + txt$
    .Left = (Screen.Width - .Width) / 2
    If (AngebotY% < 0) Then AngebotY% = (Screen.Height - .Height) / 2
    If (iNewLine = 0) Then
        .Top = AngebotY%
    End If
End With

If (AngebotInd% < 0) And (AnzLocalAngebote% = 0) And (AngebotDirektEingabe% = False) Then
    If (angLief% > 0) Then
        frmAngebote.Show
        frmAngebote.cmdAngebotNeu.Value = True
    Else
        Unload frmAngebote
        Call clsError.DefErrPop: Exit Sub
    End If
Else
    frmAngebote.Show 1
End If

'If (AngebotInd% = -1) And (angInd% >= 0) Then
'    AngebotInd% = 0
'End If

angInd% = AngebotInd%
angTemporaer = AngebotTemporaer
angManuell = AngebotManuell
If (AngebotInd% > 0) Then
    iAng% = AngebotInd% - 1
    
    If (LocalAngebotRec(iAng%).LaufNr = 2000) Then
        With clsManuelleAngebote1
            .pzn = AngebotPzn$
            
            .bm = LocalAngebotRec(iAng%).bm
            .mp = LocalAngebotRec(iAng%).mp
            .gh = LocalAngebotRec(iAng%).gh
            .st = "M"
            .zr = LocalAngebotRec(iAng%).zr
            
            If (AngebotTemporaer) Then
                angTemporaer = 0
                .Saisonal = 1
            Else
                .Saisonal = 0
            End If
            
            .BarPreis = 0
            .rest = String(Len(.rest), 0)
            
            NeuSpeichern% = True
            
            If (iAng% > 0) Then
                VglInd% = iAng% - 1
                If (LocalAngebotRec(VglInd%).LaufNr > 1000) Then
                    VglInd% = VglInd% - 1
                End If
                If (VglInd% >= 0) Then
                    iAngLen% = 43   'Len(LocalAngebotRec(iAng%))
                    sAng1$ = String(iAngLen%, 0)
                    sAng2$ = String(iAngLen%, 0)
                    CopyMemory ByVal sAng1$, LocalAngebotRec(iAng%), iAngLen%
                    CopyMemory ByVal sAng2$, LocalAngebotRec(VglInd%), iAngLen%
                    If (sAng1$ = sAng2$) Then
                        LocalAngebotRec(iAng%).LaufNr = LocalAngebotRec(VglInd%).LaufNr
                        NeuSpeichern% = False
                    End If
                End If
            End If
            If (NeuSpeichern%) And (iAng% < (AnzLocalAngebote% - 1)) Then
                VglInd% = iAng% + 1
                If (LocalAngebotRec(VglInd%).LaufNr > 1000) Then
                    VglInd% = VglInd% + 1
                End If
                If (VglInd% <= (AnzLocalAngebote% - 1)) Then
                    iAngLen% = 43   'Len(LocalAngebotRec(iAng%))
                    sAng1$ = String(iAngLen%, 0)
                    sAng2$ = String(iAngLen%, 0)
                    CopyMemory ByVal sAng1$, LocalAngebotRec(iAng%), iAngLen%
                    CopyMemory ByVal sAng2$, LocalAngebotRec(VglInd%), iAngLen%
                    If (sAng1$ = sAng2$) Then
                        LocalAngebotRec(iAng%).LaufNr = LocalAngebotRec(VglInd%).LaufNr
                        NeuSpeichern% = False
                    End If
                End If
            End If
            
            If (NeuSpeichern%) Then
                If (AngeboteDbOk) Then
                    SQLStr = "INSERT INTO Angebote (Pzn,Gh,St,Bm,Mp,Zr,Saisonal,BarPreis)"
                    SQLStr = SQLStr + " VALUES (" + .pzn + "," + CStr(.gh) + ",'" + .st + "'," + CStr(.bm) + "," + CStr(.mp) + "," + clsOpTool.uFormat(.zr, "0.00") + "," + CStr(.Saisonal) + "," + clsOpTool.uFormat(.BarPreis, "0.00")
                    SQLStr = SQLStr + ")"
                    ManuellAngeboteDB1.ActiveConn.CommandTimeout = 120
                    Call ManuellAngeboteDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                Else
                    FabsErrf% = .IndexInsert(0, AngebotPzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        .PutRecord (FabsRecno& + 1)
                    End If
                End If
                LocalAngebotRec(iAng%).LaufNr = AnzManuelle% + 1
            End If
        
            LocalAngebotRec(iAng%).IstManuell = 1
            
        End With
    End If
    
    angNm% = 0
    angBm% = LocalAngebotRec(iAng%).bm
    If (LocalAngebotRec(iAng%).st <> "P") Then
        If (LocalAngebotRec(iAng%).st <> "S") Then
            angNm% = LocalAngebotRec(iAng%).mp
        End If
        If (Para1.Land = "A") And (LocalAngebotRec(iAng%).mp2 > 0) Then angNm% = LocalAngebotRec(iAng%).mp2
    End If
    angLief% = LocalAngebotRec(iAng%).ghBest
    
    If (LocalAngebotRec(iAng%).LaufNr > 1000) Then
        AngAep# = LocalAngebotRec(iAng%).AepAng
    Else
        AngAep# = LocalAngebotRec(iAng%).AepAngOrg
    End If
    
    angZr! = LocalAngebotRec(iAng%).zr
    
    angInd% = LocalAngebotRec(iAng%).LaufNr
    If (angInd% > 1000) Then angInd% = 250
    angManuell = LocalAngebotRec(iAng%).IstManuell
End If

angY% = AngebotY%

Call clsError.DefErrPop
End Sub

Sub LöscheAngebote(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LöscheAngebote")
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
Dim i&, lRecs&

If (AngeboteDbOk) Then
    SQLStr = "DELETE * FROM Angebote WHERE Gh=" + CStr(iLief)
    ManuellAngeboteDB1.ActiveConn.CommandTimeout = 120
    Call ManuellAngeboteDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
Else
    With clsManuelleAngebote1
        For i& = 1 To (.AnzDS - 1)
            Call .GetRecord(i& + 1)
            
            If (.gh = iLief%) Then
                .st = "X"
                .PutRecord (i& + 1)
            End If
        Next i&
    End With
End If

Call clsError.DefErrPop
End Sub

Function ErstAngebot%(pzn$, angInd%, angManuell As Byte, angBm%, angNm%, angLief%, AngAep#, angZr!, angHerst$, angGhRuf$, angGhPriorität$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ErstAngebot%")
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
Dim ret%, ind%
Dim h$

ret% = False

AngebotPzn$ = pzn$
AngebotAuswahl% = False
AngebotModus% = 0
AngebotLief% = angLief%
AngebotHerst$ = angHerst$
AngebotGhRuf$ = angGhRuf$
AngebotGhPriorität$ = angGhPriorität$

AngebotBm% = angBm%
AngebotNm% = angNm%

AngebotDirektEingabe% = False

Call Angebote

If (AngebotInd% <> 0) Then
    angInd% = 2
    If (AngebotInd% < 0) Then
        ret% = True
        ind% = Abs(AngebotInd%) - 1
        angBm% = LocalAngebotRec(ind%).bm
        If (LocalAngebotRec(ind%).st = "P") Then
            angNm% = 0
        Else
            If (LocalAngebotRec(ind%).st <> "S") Then
                angNm% = LocalAngebotRec(ind%).mp
            End If
            If (Para1.Land = "A") And (LocalAngebotRec(ind%).mp2 > 0) Then angNm% = LocalAngebotRec(ind%).mp2
        End If
        angLief% = LocalAngebotRec(ind%).ghBest
'        angAep# = LocalAngebotRec(ind%).AepKalk
'weg am 13.7.2016
'        AngAep# = LocalAngebotRec(ind%).AepAngOrg
'neu am 13.7.2016
        If (LocalAngebotRec(ind%).LaufNr > 1000) Then
            AngAep# = LocalAngebotRec(ind).AepAng
        Else
            AngAep# = LocalAngebotRec(ind).AepAngOrg
        End If
        
        angZr! = LocalAngebotRec(ind%).zr
        
        ret% = LocalAngebotRec(ind%).LaufNr
        If (ret% > 1000) Then ret% = 250
        angManuell = LocalAngebotRec(ind%).IstManuell
    
    End If
Else
    angInd% = 1
End If

ErstAngebot% = ret%

Call clsError.DefErrPop
End Function

Function TesteAufPassendesAngebot%(pzn$, angBm%, angLief%, AngAep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TesteAufPassendesAngebot%")
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
Dim ret%, ind%
Dim h$

ret% = False

AngebotPzn$ = pzn$
AngebotAuswahl% = False
AngebotModus% = 99
AngebotLief% = angLief%
AngebotHerst$ = ""
AngebotGhRuf$ = ""

AngebotBm% = angBm%
AngebotNm% = 0

AngebotDirektEingabe% = False

Call Angebote

If (AngebotInd% < 0) Then
    ret% = True
    
    ind% = Abs(AngebotInd%) - 1
'    angBm% = LocalAngebotRec(ind%).bm
'    If (LocalAngebotRec(ind%).st = "P") Then
'        angNm% = 0
'    Else
'        angNm% = LocalAngebotRec(ind%).mp
'    End If
'    angLief% = LocalAngebotRec(ind%).ghBest
    
'weg am 13.7.2016
'        AngAep# = LocalAngebotRec(ind%).AepAngOrg
'neu am 13.7.2016
        If (LocalAngebotRec(ind%).LaufNr > 1000) Then
            AngAep# = LocalAngebotRec(ind).AepAng
        Else
            AngAep# = LocalAngebotRec(ind).AepAngOrg
        End If
    
'    angZr! = LocalAngebotRec(ind%).zr
    
    ret% = LocalAngebotRec(ind%).LaufNr
    If (ret% > 1000) Then ret% = 250
End If

TesteAufPassendesAngebot% = ret%

Call clsError.DefErrPop
End Function

Sub InitDirektBewertung(Optional angLief% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitDirektBewertung")
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

AngebotInd% = -1
AngebotAuswahl% = True
If (angLief% >= 0) Then AngebotActLief% = angLief%
AngebotDirektEingabe% = False
    
IstBewertung% = True
    
Call DirektBewertung(0)

Call clsError.DefErrPop
End Sub

Sub RechneDirektBewertung(pzn$, angBm%, angNm%, angZr!, angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RechneDirektBewertung")
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

AngebotPzn$ = pzn$
AngebotBm% = angBm%
AngebotNm% = angNm%
AngebotZr! = angZr!
    
AngebotBMopt! = angBMopt!
AngebotBMast% = angBMast%
AngebotWgAst% = angWgAst%
AngebotTaxeAep# = angTaxeAep#
AngebotAstAep# = angAstAep#
AngebotPznLagernd% = angPznLagernd%

Call DirektBewertung(1)

Call clsError.DefErrPop
End Sub

Function ZeigeDirektBewertung#(IstBewertungOk%, angY%, AnzeigeFlag%, Optional modus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeDirektBewertung#")
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

AngebotY% = angY%
    
Call DirektBewertung(2, AnzeigeFlag%)

IstBewertungOk% = BewertungOk%

If (modus% = 0) Then
    ZeigeDirektBewertung# = GesBewertung#
End If

Call clsError.DefErrPop
End Function

Sub ZeigeWumsatz(DarfDetails%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeWumsatz")
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
Dim iAng%

Load frmWumsatzPrognosen
frmWumsatzPrognosen.cmdF2.Enabled = DarfDetails%
frmWumsatzPrognosen.Show 1

Call clsError.DefErrPop
End Sub

Sub WinArtDebug(sDebug$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WinArtDebug")
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
Dim WINARTDEB%, ret%
Dim i%, ind%
Dim s$

'If (WinArtDebugAktiv%) Then
    s$ = Format(Now, "dd.mm.yyyy") + " " + Format(Now, "hh:nn:ss") + "   " + sDebug$
    
    WINARTDEB% = FreeFile
    Open "WINART.DEB" For Append As #WINARTDEB%
    Print #WINARTDEB%, s
    Close #WINARTDEB%
'End If

Call clsError.DefErrPop
End Sub

Sub ZeigeStatbild(SuchPzn$, pForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeStatbild")
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
Dim TaskId&, FabsRecno&
Dim i%, STATBILD%, erg%, FabsErrf%
Dim dpreis#
Dim s$, pzn$, ch$, SQLStr$
Dim recTaxe As Recordset

If (SuchPzn$ = "") Then Call clsError.DefErrPop: Exit Sub

'FabsErrf% = Ass1.IndexSearch(0, SuchPzn$, FabsRecno&)
'If (FabsErrf% = 0) Then
'    Ass1.GetRecord (FabsRecno& + 1)
'
'    FabsErrf% = Ast1.IndexSearch(0, SuchPzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        Ast1.GetRecord (FabsRecno& + 1)
'
'        For i% = 1 To 2
'            If (i% = 1) Then
'                STATBILD% = clsDat.FileOpen("statb-" + Para1.user + ".$$$", "W")
'            Else
'                STATBILD% = clsDat.FileOpen("statbild.$$$", "W")
'            End If
'
'            Ast1.PrintFile (STATBILD%)
'            Ass1.PrintFile (STATBILD%)
'    '        Put STATBILD%, , Ast1
'    '        Put STATBILD%, , Ass1
'
'            s$ = String$(200, 0)
'            Put STATBILD%, , s$
'
'            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(SuchPzn$)
'            Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
'            If (recTaxe.EOF = False) Then
'                dpreis# = recTaxe!EK / 100
'                Call DxToMBFd(dpreis#)
'                Put STATBILD%, , dpreis#
'                dpreis# = recTaxe!vk / 100
'                Call DxToMBFd(dpreis#)
'                Put STATBILD%, , dpreis#
'            End If
'
'            Put STATBILD%, , Left$(s$, 2)
'
'            Close #STATBILD%
'
'            If (Para1.Land = "D") Then
'                Exit For
'            End If
'        Next i%
'
'        TaskId& = Shell("\user\dosrun.bat " + Para1.user + " statbild.exe", vbNormalFocus)
'        Call clsOpTool.WarteAufTaskEnde(TaskId&, pForm)
''        AppActivate Me
'
'    End If
'End If
        
TaskId& = Shell("WinStat.exe " + SuchPzn$, vbNormalFocus)

Call clsError.DefErrPop
End Sub

Function WechselFenster%(zeilen$(), start%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WechselFenster%")
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
Dim i%, max%

max% = UBound(zeilen$)
ReDim WechselZeilen$(max%)
For i% = 0 To max%
    WechselZeilen$(i%) = zeilen$(i%)
Next i%
WechselStart% = start%

frmWechsel.Show 1

start% = WechselStart%
WechselFenster% = WechselErg%

Call clsError.DefErrPop
End Function

Sub LeseStiftEinlesen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LeseStiftEinlesen")
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

frmDp04.Show 1

Call clsError.DefErrPop
End Sub

Sub Stammdaten(pzn$, Optional stInd& = 0, Optional stWert As Variant = "", Optional stClass As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Stammdaten")
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
Dim x&, hWnd&, ThreadID&, TaskId&
Dim sPara$

'hWnd = GetActiveWindow
'x& = GetWindowThreadProcessId(hWnd&, ThreadID)
ThreadID = GetCurrentProcessId
hWnd = GetForegroundWindow

sPara = pzn + " " + CStr(stInd) + " " + CStr(0) + " " + CStr(hWnd)
'MsgBox (sPara)

TaskId& = Shell("WinLief.exe " + sPara$, vbNormalFocus)
Call clsOpTool.WarteAufTaskEnde(TaskId&, Screen.ActiveForm)
'AppActivate Me.Caption

'StammdatenPzn$ = pzn$
'StammdatenModus% = stInd& Mod 10
'StammdatenStart$ = Format(stInd& \ 10, "0000")
'StammdatenWert = stWert
'Set StammdatenClass = stClass
'frmLiefStammdaten.Show 1

Call clsError.DefErrPop
End Sub

Function OpenOpConnect%(fPznDb As Database, pProfileDb As Database)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenOpConnect%")
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

FremdPznTest% = True
FremdPznOk% = 0

On Error Resume Next
Set FremdPznDB = OpenDatabase("FremdPzn.mdb", False, False)
If (Err = 0) Then
    Set OpPartnerDB = OpenDatabase("Partner.mdb", False, True)
End If
If (Err = 0) Then
    FremdPznOk% = True
    Set fPznDb = FremdPznDB
    Set FremdPznRec = FremdPznDB.OpenRecordset("Artikel", dbOpenTable)
    FremdPznRec.Index = "Unique"
    
    Set pProfileDb = OpPartnerDB
    Set OpPartnerRec = OpPartnerDB.OpenRecordset("PartnerProfile", dbOpenTable)
    OpPartnerRec.Index = "Unique"
End If

OpenOpConnect% = FremdPznOk%

Call clsError.DefErrPop
End Function

Function ZeigePosPartner%(pzn$, txt$, BestellFlag%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigePosPartner%")
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
Dim ret%

ret% = 0

OpPartnerPzn$ = pzn$
OpPartnerTxt$ = txt$

If (BestellFlag% > 0) Then
    OpPartnerInitLief% = BestellFlag%
    BestellFlag% = True
Else
    OpPartnerInitLief% = 0
End If

OpPartnerBestell% = BestellFlag%

frmPosPartner.Show 1
If (EditErg%) Then
    ret% = Val(EditTxt$)
End If

ZeigePosPartner% = ret%

Call clsError.DefErrPop
End Function

Function HoleBenutzerSignatur%(Optional iModus% = 3)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleBenutzerSignatur%")
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
Dim ret%, RFID%
Dim l&
Dim h$

SignaturEingabeModus = iModus
SignaturEingabeErg = 0

RFID = 0
If ((SignaturEingabeModus And 2)) Then
    h$ = Space$(100)
    l& = GetPrivateProfileString("Allgemein", "RFID", h, h, 101, CurDir + "\dp.ini")
    h$ = Trim$(Left$(h$, l&))
    If (Len(h$) > 0) Then
        h = "," + h + ","
        RFID = (InStr(h, "," + CStr(Val(Para1.user)) + ",") > 0)
    End If
End If
            
If (RFID) Then
    'Call OpChipAufruf
    Dim TaskId&
    TaskId = clsOpTool.GetTaskId("Benutzer Signatur")
    If (TaskId = 0) Then
        TaskId = Shell("OpChip VK", vbNormalFocus)
    End If

    frmRFID.Show 1
Else
    frmSignatur.Show 1
End If
HoleBenutzerSignatur% = SignaturEingabeErg%

Call clsError.DefErrPop
End Function

Sub OpChipAufruf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpChipAufruf")
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
'Dim pzn As Long
Dim s As String, SQL As String, zeile As String, CheckGHFile As String, sAction$
Dim i, fhtmp As Integer
Dim erfolg As Boolean
ReDim argv(0) As Variant
Dim fhtmpname As String
Dim StartTimer As Single

fhtmpname = "AE" + Format(Now, "nnss") + CStr(Val(Para1.user)) + "L" + ".$$$"
On Error Resume Next
If Dir(CurDir + "\" + fhtmpname) > "" Then Kill CurDir + "\" + fhtmpname
CheckGHFile = CurDir + "\" + fhtmpname

Shell CurDir + "\OpChip.EXE " + fhtmpname, vbNormalFocus

'InCheckSecPh = True
StartTimer = Timer

Do
  If Dir(CurDir + "\" + fhtmpname) > "" Then
    If GetFileSize(CurDir + "\" + fhtmpname) > 0 Then
      Exit Do
    End If
  End If
  
'  If ((Timer - StartTimer) > 15) Then
'    GHAbbruch = True
'    Exit Do
'  End If
Loop
On Error GoTo DefErr

'InCheckSecPh = False

On Error Resume Next
erfolg = False

'If (GHAbbruch) Then
'Else
  fhtmp = FreeFile
  Open CurDir + "\" + fhtmpname For Input As #fhtmp
  Line Input #fhtmp, s
  Close #fhtmp
  
  i = InStr(s, ",")
  If (i > 0) Then
    SignaturEingabeErg = Val(Left(s, i - 1))
  End If
'End If

Call clsError.DefErrPop
End Sub

Function GetFileSize(ByVal strFilename As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function



