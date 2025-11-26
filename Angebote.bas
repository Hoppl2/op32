Attribute VB_Name = "modAngebote"
Option Explicit

Type LocalAngeboteStruct
    gh As Byte
    st As String * 1
    bm As Integer   'Byte
    mp As Long
    zr As Single
    bmOrg As Integer    'Byte
    mpOrg As Long
    AepKalk As Double
    AepAng As Double
    AepAngOrg As Double
    IstManuell As Byte
    LaufNr As Integer
    recno As Long
    ghBest As Byte
    Saisonal As Byte
    IstMvda As Byte
    mp2 As Long
End Type


Public AngebotPzn$
Public AngebotModus%
Public AngebotInd%
Public AngebotAuswahl%
Public AngebotAep#
Public AngebotLief%
Public AngebotBm%
Public AngebotNm%
Public AngebotY%
Public AngebotActLief%
Public AngebotNeuLief%
Public AngebotGhBest%
Public AngebotZr!
Public AngebotTemporaer As Byte
Public AngebotManuell As Byte
Public AngebotMitZr As Byte
Public AngebotBMopt!
Public AngebotBMast%
Public AngebotWgAst%
Public AngebotTaxeAep#
Public AngebotAstAep#
Public AngebotPznLagernd%
Public AngebotHerst$
Public AngebotLiefsFuerHerst$
Public AngebotGhRuf$
Public AngebotGhPriorität$
Public AngebotDirektEingabe%

Public AngebotEditBm%, AngebotEditNm%
Public AngebotEditZr!
Public AngebotEditRecno&

Public AngebotNeu%

Public GesBewertung#
Public BewertungOk%

Dim zr!, fr!
Dim bm%, nr%, OrgBm%
Dim gspart#, prozgspart#
Dim OrgMp&

Public LocalAngebotRec(50) As LocalAngeboteStruct
Public GhAngebotMax&
Public AEPorg#
Public TaxeHAP#

Public AnzLocalAngebote%


Dim NrWert#, ZrWert#, FrWert#, BarRabWert#, LaKo#, PeKo#, Manko#, StaffelAb#
Dim BMopt!, BMoptOrg!, BMast%, WgAst%

Dim FabsRecno&
Dim FabsErrf%

Dim GhAnfSatz&, ManAnfSatz&

Dim GesBm%, GesNr%
Dim GesAepOrg#, GesAepScreen#, GesNrWert#, GesBarRabWert#, GesLaKo#, GesPeKo#, GesStaffelAb#, GesAepKalk#
Dim GesGspart#, GesProzGspart#, GesBarRabScreen#

Public IstBewertung%
Public AnzManuelle%

Dim AngeboteDB As Database
Dim AngeboteDaoRec As Recordset
Dim AngeboteOk%
Dim SQLStr$

Private Const DefErrModul = "ANGEBOTE.BAS"

Sub Angebote()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Angebote")
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
Dim AngebotDa%, ErrNumber%
Dim TaxeAep#, asatz&, SQLStr$, DBName$
Dim pzn$

pzn$ = AngebotPzn$

If (Para1.Land = "A") Then
    AngeboteOk = 0
    DBName$ = "Angebote.mdb"
    If (Dir(DBName) <> "") Then
        On Error Resume Next
        Err.Clear
        Set AngeboteDB = OpenDatabase(DBName$, False, False)
        ErrNumber% = Err.Number
        On Error GoTo DefErr
    
        AngeboteOk = (ErrNumber% = 0)
    End If
End If

AngebotDa% = AngebotCheck%(pzn$)

If (AngebotAuswahl% = False) Then
    AngebotInd% = Abs(AngebotDa%)
    If (AngebotBm% = 0) Then Call clsError.DefErrPop: Exit Sub
End If
    
If (AngebotDa% Or AngebotAuswahl%) Then
    asatz& = 0
    Ass1.pzn = String$(7, 0)
    Ast1.pzn = String$(7, 0)
    If (ArtikelDBok) Then
        SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
        If (FabsErrf% = 0) Then
            asatz& = 1
        End If
    Else
        FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Ass1.GetRecord (FabsRecno& + 1)
        End If
    
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            asatz& = FabsRecno&
            Ast1.GetRecord (FabsRecno& + 1)
        End If
    End If
  
    If AngebotAuswahl% Or (AngebotAuswahl% = 0 And AngebotLief% = 0) Or (AngebotAuswahl% = 0 And AngebotModus% = 99) Then
       TaxeAep# = 0#
       TaxeHAP# = 0#

        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
        If (TaxeAdoDBok) Then
            On Error Resume Next
            TaxeAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
            If (TaxeAdoRec.EOF = False) Then
                TaxeAep# = TaxeAdoRec!EK / 100
                TaxeHAP# = TaxeAdoRec!HAP / 100
                If (asatz& <= 0) Then
                    Call clsOpTool.Taxe2ast(pzn$)
                End If
                WgAst% = Val(Ast1.wg)
            End If
'            TaxeAdoRec.Close
        Else
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                TaxeAep# = TaxeRec!EK / 100
                TaxeHAP# = TaxeRec!HAP / 100
                If (asatz& <= 0) Then
                    Call clsOpTool.Taxe2ast(pzn$)
                End If
                WgAst% = Val(Ast1.wg)
            End If
        End If

        Call AngebotSelect(pzn$, TaxeAep#, (Ast1.pzn <> String$(7, 0)))
    End If
End If

If (AngeboteOk) Then
    AngeboteDB.Close
End If

Call clsError.DefErrPop
End Sub

Function AngebotCheck%(SuchPzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotCheck%")
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
Dim von&, bis&, Index&, Such$, satz&
Dim s$, h2$

GhAnfSatz& = -1&
ManAnfSatz& = -1&

ret% = False
If (Val(SuchPzn$) = 9999999) Then Call clsError.DefErrPop: Exit Function

If (Para1.Land = "A") Then
    If (AngeboteOk) Then
        SQLStr = "SELECT * FROM Angebote WHERE Pzn=" + SuchPzn
        SQLStr = SQLStr + " ORDER BY Gh,KonditionsArt DESC,StaffelMenge ASC"
        Set AngeboteDaoRec = AngeboteDB.OpenRecordset(SQLStr)
        ret = Not (AngeboteDaoRec.EOF)
    End If
    AngebotCheck = ret
    Call clsError.DefErrPop: Exit Function
End If

If (AngeboteDbOk) Then
    On Error Resume Next
    AngeboteRec.Close
    On Error GoTo DefErr
    SQLStr = "SELECT * FROM Angebote WHERE Pzn=" + SuchPzn
    SQLStr = SQLStr + " ORDER BY Gh,Bm"
    AngeboteRec.Open SQLStr, AngeboteDB1.ActiveConn
    ret = Not (AngeboteRec.EOF)

    On Error Resume Next
    ManuelleAngeboteRec.Close
    On Error GoTo DefErr
    SQLStr = "SELECT * FROM Angebote WHERE Pzn=" + SuchPzn
    SQLStr = SQLStr + " ORDER BY Gh,Bm"
    ManuelleAngeboteRec.Open SQLStr, ManuellAngeboteDB1.ActiveConn
Else
    von& = 1&
    bis& = clsAngebote1.AnzDS - 1
    'Such$ = Format(SuchPzn&, "0000000")
    
    Do While (von& <= bis&)
        Index& = (von& + bis&) \ 2
        Call clsAngebote1.GetRecord(Index& + 1)
        s$ = clsAngebote1.pzn
        If (SuchPzn$ = s$) Then
            GhAnfSatz& = Index&
            Exit Do
        ElseIf (SuchPzn$ < s$) Then
            bis& = Index& - 1
        Else
            von& = Index& + 1
        End If
    Loop
    
    If (GhAnfSatz& >= 1) Then
        Do
            GhAnfSatz& = GhAnfSatz& - 1
            Call clsAngebote1.GetRecord(GhAnfSatz& + 1)
            s$ = clsAngebote1.pzn
            If (SuchPzn$ <> s$) Then
                Exit Do
            End If
        Loop
        GhAnfSatz& = GhAnfSatz& + 1
        ret% = True
    End If

    'If (clsManuelleAngebote1.IndexSearch(0, SuchPzn$, ManAnfSatz&) = 0) Then ret% = True
    
    If (clsManuelleAngebote1.IndexSearch(0, SuchPzn$, satz&) = 0) Then
        Do While (satz& > 0) And (satz& < clsManuelleAngebote1.AnzDS)
            
            Call clsManuelleAngebote1.GetRecord(satz& + 1)
            If (SuchPzn$ <> clsManuelleAngebote1.pzn) Then Exit Do
            
            If (clsManuelleAngebote1.st <> "X") Then
                ret% = True
                ManAnfSatz& = satz&
                Exit Do
            End If
            
            If (clsManuelleAngebote1.IndexNext(0, satz&, SuchPzn$, satz&) <> 0) Then satz& = 0
        Loop
    End If
End If


AngebotLiefsFuerHerst$ = ""

If (AngebotBm% > 0) Then
    h2$ = AngebotHerst$
    ind% = InStr(h2$, vbCr)
    If (ind% > 0) Then
        AngebotLiefsFuerHerst$ = Mid$(h2$, ind% + 1) + ","
        h2$ = Left$(h2$, ind% - 1)
    End If

'    AngebotLiefsFuerHerst$ = LifZus1.GetLiefFuerHerst(UCase(AngebotHerst$), AngebotLiefsFuerHerst$)
End If

If (AngebotLiefsFuerHerst$ <> "") Then ret% = True

AngebotCheck% = ret%

Call clsError.DefErrPop
End Function

Sub AngebotSelect(SuchPzn$, taxAEP#, PznLagernd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotSelect")
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
Dim i%, j%, k%, bmOrg%, bm1%, gh%, iCol%, ind%, ind2%, GhOk%, GhPrio1%, GhPrio2%, MinGhPrio%, raus%
Dim BmOptSelect!
Dim AepOrg2#, AepKalk#, minaep#, aep1#, aep2#, BarPreis#
Dim satz&, mp&, mp2&, mp1&, mpOrg&, MaxBm&
Dim pzn$, st$, h$, h2$, h3$
Dim iLocalAngebotRec As LocalAngeboteStruct
Dim AngebotsClass As Object
Dim LaufNr As Byte
Dim Saisonal As Byte

'BMoptOrg! = BMopt!
BmOptSelect! = BMopt!
If (AngebotModus% <> 2) Then
    If (Ass1.pzn <> String$(7, 0)) Then
        BMopt! = Ass1.opt
        
        If (Ass1.vbm > 0) Then
            BmOptSelect! = Ass1.vbm
        Else
            BmOptSelect! = Ass1.opt
        End If
        If (AngebotBm% > BmOptSelect!) Then
            BmOptSelect! = AngebotBm%
        End If
        BMast% = Ass1.bm
    Else
        BMopt! = 0
        BMast% = 0
        BmOptSelect! = 0
    End If
End If
BMoptOrg! = BMopt!

'If (Ast1.pzn <> String$(7, 0)) Then
If (PznLagernd%) Then
'    AEPorg# = Ast1.aep
    AEPorg# = taxAEP#
    If (AngebotModus% <> 2) Then
        AepOrg2# = Ast1.aep
    Else
        AepOrg2# = AngebotAstAep#
    End If
Else
    AEPorg# = 0#
    AepOrg2# = 0#
End If


If BMopt! > 3200 Then BMopt! = 3200
If BMopt! <= 0 Then BMopt! = 0.1 ' 1

If BmOptSelect! > 3200 Then BmOptSelect! = 3200
If BmOptSelect! <= 0 Then BmOptSelect! = 0.1 ' 1

AnzLocalAngebote% = 0

If (AngebotModus% = 2) Then
    LocalAngebotRec(0).gh = AngebotActLief%
    LocalAngebotRec(0).st = "M"
    LocalAngebotRec(0).bm = AngebotBm%
    LocalAngebotRec(0).mp = AngebotNm%
    LocalAngebotRec(0).zr = AngebotZr!
    LocalAngebotRec(0).bmOrg = AngebotBm%
    LocalAngebotRec(0).mpOrg = AngebotNm%
    LocalAngebotRec(0).IstManuell = 0
    LocalAngebotRec(0).ghBest = AngebotActLief%
    LocalAngebotRec(0).Saisonal = 0
    LocalAngebotRec(0).AepAngOrg = 0
    LocalAngebotRec(0).IstMvda = 0
    AnzLocalAngebote% = 1
    
    AepKalk# = AngebotAuswerten#(0)
    If (IstBewertung% = False) Then
        Call AngebotZeigen(0, AepKalk#, frmAngebote.flxAngebote)
    
        With frmAngebote.flxAngebote
            .Cols = AnzLocalAngebote% + 1
            .col = 1
            .row = 1
            .RowSel = .Rows - 1
        End With
    
        With frmAngebote
            .lblAngeboteWert(0) = Format(BMopt!, "0.0")
    
'            If (Ast1.pzn <> String$(7, 0)) Then
'            If (PznLagernd%) Then
'                AepOrg2# = Ast1.aep
'            Else
'                AepOrg2# = 0#
'            End If
            .lblAngeboteWert(1) = Format(AepOrg2#, "0.00")
            
            .lblAngeboteWert(2) = Format(taxAEP#, "0.00")
        End With
    End If

    Call clsError.DefErrPop: Exit Sub
End If

h2$ = AngebotLiefsFuerHerst$
h3$ = AngebotGhRuf$

AnzManuelle% = 0
If (Para1.Land = "A") Then
    If (AngeboteOk) Then
        AngeboteDaoRec.MoveFirst
        
        LaufNr = 1
        
        Do
            If (AngeboteDaoRec.EOF) Then
                Exit Do
            End If
            
            gh% = AngeboteDaoRec!gh
            st$ = Trim(AngeboteDaoRec!KonditionsArt)
            
            h$ = Format(gh%, "000") + ","
            ind% = InStr(h3$, h$)
            If (ind% > 0) Then h3$ = Left$(h3$, ind% - 1) + Mid$(h3$, ind% + 4)
            
            Saisonal = 0
            BarPreis# = 0
            
            bm% = AngeboteDaoRec!StaffelMenge
            
            mp& = AngeboteDaoRec!Kondition
            If (st$ = "SP") Then
            ElseIf (st$ = "GR") Then
                If (AEPorg# > 0.009) Then
                    mp = AEPorg# * (100 - mp / 100) '/ 100
                End If
                st$ = "SP"
            Else
                mp& = AngeboteDaoRec!Kondition / 100
            End If
            mp2 = 0
            
            zr! = 0!
            bmOrg% = bm%: mpOrg& = mp&
                
            If (st$ <> "NRS") Then
                If (bm% < BmOptSelect!) Then
                    For i = (AnzLocalAngebote - 1) To 0 Step -1
                        raus = 0
                        If (LocalAngebotRec(i).gh = gh%) And (LocalAngebotRec(i).st = st$) Then
                            If (LocalAngebotRec(i).bm < bm%) Then
                                raus = True
                            End If
                        End If
                        If (raus) Then
                            AnzLocalAngebote = AnzLocalAngebote - 1
                            LaufNr = LaufNr - 1
                        Else
                            Exit For
                        End If
                    Next i
                    bm% = Int(BmOptSelect!)
                End If
            End If
            
            If (st$ = "SP") Then
                If (AngeboteDaoRec!Deckelung) Then
                    mp = (mp * bmOrg + AEPorg * 100 * (bm - bmOrg)) / bm
                End If
            ElseIf (st$ = "NRP") Then
                '* Naturalrabatt Stück
                mp = bm * (mp / 100#)
            ElseIf (st$ = "NRS") Then
                '* Naturalrabatt Prozentuell
                If (bm% < BmOptSelect!) Then
                    bm% = Int(BmOptSelect!)
                End If
                If (AngeboteDaoRec!Deckelung = 0) Then
                    bm1 = bm    'BmOptSelect
                    mp1 = 0
                    While (bm1 >= bmOrg)
                        bm1 = bm1 - bmOrg
                        mp1 = mp1 + mp
                    Wend
                    If (bm1 > 0) Then
                        For i = (AnzLocalAngebote - 1) To 0 Step -1
                            If (LocalAngebotRec(i).gh = gh%) And (LocalAngebotRec(i).st = "N") Then
                                While (bm1 >= LocalAngebotRec(i).bmOrg)
                                    bm1 = bm1 - LocalAngebotRec(i).bmOrg
                                    mp1 = mp1 + LocalAngebotRec(i).mpOrg
                                Wend
                            End If
                        Next i
                    End If
                    mp = mp1
                End If
                For i = (AnzLocalAngebote - 1) To 0 Step -1
                    If (LocalAngebotRec(i).gh = gh%) And (LocalAngebotRec(i).st = "S") And (LocalAngebotRec(i).mp2 = 0) Then
                        If (bm >= LocalAngebotRec(i).bmOrg) Then
                            mp2 = mp
                            mp = LocalAngebotRec(i).mpOrg
                            st = "SP"
                            Exit For
                        End If
                    End If
                Next i
            End If
            
            '* Naturalrabatt
    '        If (st$ = "NRP" Or st$ = "NRS") And bm% > 0 And mp& > 0 Then
    '            'kleinstes Angebot unterhalb der optimalen Bestellmenge
    '            While bm% < BmOptSelect!
    '              bm% = bm% + bmOrg%
    '              mp& = mp& + mpOrg&
    '            Wend
    '            If bm% <> bmOrg% And bm% > BmOptSelect! Then
    '              'hoppla; eins zu weit, da erst bei größer gleich abgebrochen wird
    '              bm% = bm% - bmOrg%
    '              mp& = mp& - mpOrg&
    '            End If
    '        End If
        
            '* Preisrabatt
    '        If (st$ = "SP" Or st$ = "GR") And bm% > 0 Then
    '            If (bm% < BmOptSelect!) Then
    '                bm% = Int(BmOptSelect!)
    '            End If
    '        End If
            
            If (st$ = "NRP") Then
                st$ = "%"
            End If
        
            While bm% > 0 And mp& >= 0
                If AnzLocalAngebote% < 50 Then
                    LocalAngebotRec(AnzLocalAngebote%).gh = gh%
                    LocalAngebotRec(AnzLocalAngebote%).st = st$
                    LocalAngebotRec(AnzLocalAngebote%).bm = bm%
                    LocalAngebotRec(AnzLocalAngebote%).mp = mp&
                    LocalAngebotRec(AnzLocalAngebote%).zr = zr!
                    LocalAngebotRec(AnzLocalAngebote%).bmOrg = bmOrg%
                    LocalAngebotRec(AnzLocalAngebote%).mpOrg = mpOrg&
                    LocalAngebotRec(AnzLocalAngebote%).IstManuell = 0
                    LocalAngebotRec(AnzLocalAngebote%).LaufNr = LaufNr
                    LocalAngebotRec(AnzLocalAngebote%).recno = satz&
                    LocalAngebotRec(AnzLocalAngebote%).ghBest = gh%
                    LocalAngebotRec(AnzLocalAngebote%).Saisonal = Saisonal
                    LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = BarPreis#
                    LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
                    LocalAngebotRec(AnzLocalAngebote%).mp2 = mp2
                    If (j% = 1) Then AnzManuelle% = AnzManuelle% + 1
                    AnzLocalAngebote% = AnzLocalAngebote% + 1
                    LaufNr = LaufNr + 1
                End If
                If st$ = "M" And bm% < BmOptSelect! Then
                    'wenn das Mengenangebot kleiner ist als die optimale
                    'Bestellmenge versuchen wir noch das Angebot darüber
                    bm% = bm% + bmOrg%
                    mp& = mp& + mpOrg&
                Else
                    bm% = 0: mp& = -1
                End If
            Wend
            
            AngeboteDaoRec.MoveNext
        Loop
    End If
    
Else
    If (AngeboteDbOk) Then
        ind2 = Abs(ManuelleAngeboteAktiv)
        For j% = 0 To ind2
            If (j = 0) Then
                If Not (AngeboteRec.EOF) Then
                    AngeboteRec.MoveFirst
                End If
            ElseIf (ManuelleAngeboteRec.EOF) Then
                Exit For
            Else
                ManuelleAngeboteRec.MoveFirst
            End If
        
            LaufNr = 1
        
            Do
                If (j = 0) Then
                    If (AngeboteRec.EOF) Then
                        Exit Do
                    End If
                    
                    satz = AngeboteRec!Id
                    gh% = AngeboteRec!gh
                    st$ = Trim(AngeboteRec!st)
                Else
                    If (ManuelleAngeboteRec.EOF) Then
                        Exit Do
                    End If
                    
                    satz = ManuelleAngeboteRec!Id
                    gh% = ManuelleAngeboteRec!gh
                    st$ = Trim(ManuelleAngeboteRec!st)
                End If
                
                h$ = Format(gh%, "000") + ","
                ind% = InStr(h3$, h$)
                If (ind% > 0) Then h3$ = Left$(h3$, ind% - 1) + Mid$(h3$, ind% + 4)
                
                If (st$ = "X") Then
                    If (j% = 1) Then AnzManuelle% = AnzManuelle% + 1
                    LaufNr = LaufNr + 1
                Else
                    Saisonal = 0
                    BarPreis# = 0
                    
                    If (j% = 0) Then
                        bm% = AngeboteRec!bm
                        mp& = AngeboteRec!mp
                        zr! = 0!
                    Else
                        bm% = ManuelleAngeboteRec!bm
                        mp& = ManuelleAngeboteRec!mp
                        zr! = ManuelleAngeboteRec!zr
                        BarPreis# = ManuelleAngeboteRec!BarPreis
                        
                        If (ManuelleAngeboteRec!Saisonal) Then
                            Saisonal = 1
                        End If
                    End If
                    bmOrg% = bm%: mpOrg& = mp&
                            
                    '* Naturalrabatt
                    If st$ = "M" And bm% > 0 And mp& > 0 Then
                        'kleinstes Angebot unterhalb der optimalen Bestellmenge
                        While bm% < BmOptSelect!
                          bm% = bm% + bmOrg%
                          mp& = mp& + mpOrg&
                        Wend
                        If bm% <> bmOrg% And bm% > BmOptSelect! Then
                          'hoppla; eins zu weit, da erst bei größer gleich abgebrochen wird
                          bm% = bm% - bmOrg%
                          mp& = mp& - mpOrg&
                        End If
                    End If
            
                    '* Preisrabatt
    '                If st$ = "P" And bm% > 0 And mp& > 0 Then
                    If st$ = "P" And bm% > 0 Then
                        If bm% < BmOptSelect! Then bm% = Int(BmOptSelect!)
                    End If
                
                    While bm% > 0 And mp& >= 0
                        If AnzLocalAngebote% < 50 Then
                            LocalAngebotRec(AnzLocalAngebote%).gh = gh%
                            LocalAngebotRec(AnzLocalAngebote%).st = st$
                            LocalAngebotRec(AnzLocalAngebote%).bm = bm%
                            LocalAngebotRec(AnzLocalAngebote%).mp = mp&
                            LocalAngebotRec(AnzLocalAngebote%).zr = zr!
                            LocalAngebotRec(AnzLocalAngebote%).bmOrg = bmOrg%
                            LocalAngebotRec(AnzLocalAngebote%).mpOrg = mpOrg&
                            LocalAngebotRec(AnzLocalAngebote%).IstManuell = j
                            LocalAngebotRec(AnzLocalAngebote%).LaufNr = LaufNr
                            LocalAngebotRec(AnzLocalAngebote%).recno = satz&
                            LocalAngebotRec(AnzLocalAngebote%).ghBest = gh%
                            LocalAngebotRec(AnzLocalAngebote%).Saisonal = Saisonal
                            LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = BarPreis#
                            LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
    '                        LocalAngebotRec(AnzLocalAngebote%).mp2 = mp2
    '                        If (j% = 1) Then AnzManuelle% = AnzManuelle% + 1
                            AnzLocalAngebote% = AnzLocalAngebote% + 1
                            LaufNr = LaufNr + 1
                        End If
                        If st$ = "M" And bm% < BmOptSelect! Then
                            'wenn das Mengenangebot kleiner ist als die optimale
                            'Bestellmenge versuchen wir noch das Angebot darüber
                            bm% = bm% + bmOrg%
                            mp& = mp& + mpOrg&
                        Else
                            bm% = 0: mp& = -1
                        End If
                    Wend
                End If
                
                If (j = 0) Then
                    AngeboteRec.MoveNext
                Else
                    ManuelleAngeboteRec.MoveNext
                End If
            Loop
        Next
    Else
        ind2 = Abs(ManuelleAngeboteAktiv)
        For j% = 0 To ind2
            If (j% = 0) Then
                satz& = GhAnfSatz&
                Set AngebotsClass = clsAngebote1
            Else
                satz& = ManAnfSatz&
                Set AngebotsClass = clsManuelleAngebote1
            End If
            LaufNr = 1
            
            Do While (satz& > 0) And (satz& < AngebotsClass.AnzDS)
                
                Call AngebotsClass.GetRecord(satz& + 1)
                If (SuchPzn$ <> AngebotsClass.pzn) Then Exit Do
                
                gh% = AngebotsClass.gh
                st$ = AngebotsClass.st
                
                h$ = Format(gh%, "000") + ","
                ind% = InStr(h3$, h$)
                If (ind% > 0) Then h3$ = Left$(h3$, ind% - 1) + Mid$(h3$, ind% + 4)
                
                If (st$ = "X") Then
                    If (j% = 1) Then AnzManuelle% = AnzManuelle% + 1
                    LaufNr = LaufNr + 1
                Else
                    For i% = 0 To 4
                        If (j% = 1) And (i% = 1) Then Exit For
                        
                        Saisonal = 0
                        BarPreis# = 0
                        
                        If (j% = 0) Then
                            bm% = AngebotsClass.bm(i%)
                            mp& = AngebotsClass.mp(i%)
                            zr! = 0!
                        Else
                            bm% = AngebotsClass.bm
                            mp& = AngebotsClass.mp
                            zr! = AngebotsClass.zr
                            BarPreis# = AngebotsClass.BarPreis
                            
                            If (AngebotsClass.Saisonal) Then
                                Saisonal = 1
                            End If
                        End If
                        bmOrg% = bm%: mpOrg& = mp&
                        
                        '* Naturalrabatt
                        If st$ = "M" And bm% > 0 And mp& > 0 Then
                            'kleinstes Angebot unterhalb der optimalen Bestellmenge
                            While bm% < BmOptSelect!
                              bm% = bm% + bmOrg%
                              mp& = mp& + mpOrg&
                            Wend
                            If bm% <> bmOrg% And bm% > BmOptSelect! Then
                              'hoppla; eins zu weit, da erst bei größer gleich abgebrochen wird
                              bm% = bm% - bmOrg%
                              mp& = mp& - mpOrg&
                            End If
                        End If
                
                        '* Preisrabatt
        '                If st$ = "P" And bm% > 0 And mp& > 0 Then
                        If st$ = "P" And bm% > 0 Then
                            If bm% < BmOptSelect! Then bm% = Int(BmOptSelect!)
                        End If
                
                        While bm% > 0 And mp& >= 0
                            If AnzLocalAngebote% < 50 Then
                                LocalAngebotRec(AnzLocalAngebote%).gh = gh%
                                LocalAngebotRec(AnzLocalAngebote%).st = st$
                                LocalAngebotRec(AnzLocalAngebote%).bm = bm%
                                LocalAngebotRec(AnzLocalAngebote%).mp = mp&
                                LocalAngebotRec(AnzLocalAngebote%).zr = zr!
                                LocalAngebotRec(AnzLocalAngebote%).bmOrg = bmOrg%
                                LocalAngebotRec(AnzLocalAngebote%).mpOrg = mpOrg&
                                LocalAngebotRec(AnzLocalAngebote%).IstManuell = j%
                                LocalAngebotRec(AnzLocalAngebote%).LaufNr = LaufNr
                                LocalAngebotRec(AnzLocalAngebote%).recno = satz&
                                LocalAngebotRec(AnzLocalAngebote%).ghBest = gh%
                                LocalAngebotRec(AnzLocalAngebote%).Saisonal = Saisonal
                                LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = BarPreis#
                                LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
                                If (j% = 1) Then AnzManuelle% = AnzManuelle% + 1
                                AnzLocalAngebote% = AnzLocalAngebote% + 1
                                LaufNr = LaufNr + 1
                            End If
                            If st$ = "M" And bm% < BmOptSelect! Then
                                'wenn das Mengenangebot kleiner ist als die optimale
                                'Bestellmenge versuchen wir noch das Angebot darüber
                                bm% = bm% + bmOrg%
                                mp& = mp& + mpOrg&
                            Else
                                bm% = 0: mp& = -1
                            End If
                        Wend
                    Next i%
                End If
            
                If (j% = 0) Then
                    satz& = satz& + 1
                Else
                    If (AngebotsClass.IndexNext(0, satz&, SuchPzn$, satz&) <> 0) Then satz& = 0
                End If
            Loop
        Next j%
    End If
    
    ''''''''''''''''''''''''''
    Do
    '    If (ww.bm = 0) Then Exit Do
        If (h2$ = "") Then Exit Do
        ind% = InStr(h2$, ",")
        If (ind% > 0) Then
            gh% = Val(Left$(h2$, ind% - 1))
            
            h$ = Format(gh%, "000") + ","
            ind2% = InStr(h3$, h$)
            If (ind2% > 0) Then h3$ = Left$(h3$, ind2% - 1) + Mid$(h3$, ind2% + 4)
            
            LocalAngebotRec(AnzLocalAngebote%).gh = gh%
            LocalAngebotRec(AnzLocalAngebote%).st = "M"
            LocalAngebotRec(AnzLocalAngebote%).bm = AngebotBm%
            LocalAngebotRec(AnzLocalAngebote%).mp = AngebotNm%
            LocalAngebotRec(AnzLocalAngebote%).zr = -123.45!
            LocalAngebotRec(AnzLocalAngebote%).bmOrg = AngebotBm%
            LocalAngebotRec(AnzLocalAngebote%).mpOrg = AngebotNm%
            LocalAngebotRec(AnzLocalAngebote%).IstManuell = 0
            LocalAngebotRec(AnzLocalAngebote%).LaufNr = 1000 + gh%
    '        LocalAngebotRec(AnzLocalAngebote%).recno = satz&
            LocalAngebotRec(AnzLocalAngebote%).ghBest = gh%
            LocalAngebotRec(AnzLocalAngebote%).Saisonal = 0
            LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = 0
            LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
            AnzLocalAngebote% = AnzLocalAngebote% + 1
    
            h2$ = Mid$(h2$, ind% + 1)
        Else
            Exit Do
        End If
    Loop
    '''''''''''''''''''''''''''''
    Do
        ind% = InStr(h3$, ",")
        If (ind% > 0) Then
            gh% = Left(h3$, ind% - 1)
            
            GhOk% = True
            If (Ast1.lic = "D") And (gh% > 0) And (gh% <= LifZus1.AnzRec) Then
                LifZus1.GetRecord (gh% + 1)
                GhOk% = LifZus1.IstDirektLieferant
            End If
            
            If (GhOk%) Then
                LocalAngebotRec(AnzLocalAngebote%).gh = gh%
                LocalAngebotRec(AnzLocalAngebote%).st = "P"
                LocalAngebotRec(AnzLocalAngebote%).bm = AngebotBm%
                LocalAngebotRec(AnzLocalAngebote%).mp = AEPorg# * 100
                LocalAngebotRec(AnzLocalAngebote%).zr = 0
                LocalAngebotRec(AnzLocalAngebote%).bmOrg = AngebotBm%
                LocalAngebotRec(AnzLocalAngebote%).mpOrg = AEPorg# * 100
                LocalAngebotRec(AnzLocalAngebote%).IstManuell = 0
                LocalAngebotRec(AnzLocalAngebote%).LaufNr = 1000 + gh%
        '        LocalAngebotRec(AnzLocalAngebote%).recno = satz&
                LocalAngebotRec(AnzLocalAngebote%).ghBest = gh%
                LocalAngebotRec(AnzLocalAngebote%).Saisonal = 0
                LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = 0
                LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
                AnzLocalAngebote% = AnzLocalAngebote% + 1
            End If
            
            h3$ = Mid$(h3, ind% + 1)
        Else
            Exit Do
        End If
    Loop
    '''''''''''''''''''''''''''''
    If (AngebotDirektEingabe%) Then
        LocalAngebotRec(AnzLocalAngebote%).gh = AngebotActLief%
        LocalAngebotRec(AnzLocalAngebote%).st = "M"
        LocalAngebotRec(AnzLocalAngebote%).bm = AngebotBm%
        LocalAngebotRec(AnzLocalAngebote%).mp = AngebotNm%
        LocalAngebotRec(AnzLocalAngebote%).zr = AngebotZr!
        LocalAngebotRec(AnzLocalAngebote%).bmOrg = AngebotBm%
        LocalAngebotRec(AnzLocalAngebote%).mpOrg = AngebotNm%
        LocalAngebotRec(AnzLocalAngebote%).IstManuell = 0
        LocalAngebotRec(AnzLocalAngebote%).LaufNr = 2000
    '        LocalAngebotRec(AnzLocalAngebote%).recno = satz&
        LocalAngebotRec(AnzLocalAngebote%).ghBest = AngebotActLief%
        LocalAngebotRec(AnzLocalAngebote%).Saisonal = 0
        LocalAngebotRec(AnzLocalAngebote%).AepAngOrg = 0
        LocalAngebotRec(AnzLocalAngebote%).IstMvda = 0
        AnzLocalAngebote% = AnzLocalAngebote% + 1
    End If
    '''''''''''''''''''''''''''''
End If

minaep# = 9999999999#
MinGhPrio% = 999
MaxBm& = 0
For i% = 0 To (AnzLocalAngebote% - 1)
    AepKalk# = AngebotAuswerten#(i%)
    GhPrio1% = GetGhPriorität((LocalAngebotRec(i%).gh))
    If (AngebotAuswahl% = False) Then
        If (AngebotModus% = 0) Then
            If AepKalk# > 0 And ((AepKalk# < minaep#) Or ((AepKalk# = minaep#) And (GhPrio1% < MinGhPrio%))) And AepKalk# < AEPorg# And BMast% > 0 Then
                minaep# = AepKalk#
                MinGhPrio% = GhPrio1%
                AngebotInd% = (i% + 1) * (-1)
                AngebotLief% = LocalAngebotRec(i%).ghBest
                AngebotBm% = LocalAngebotRec(i%).bm
                AngebotNm% = 0
                If LocalAngebotRec(i%).st = "M" Then AngebotNm% = LocalAngebotRec(i%).mp
                If (Para1.Land = "A") And (LocalAngebotRec(i%).mp2 > 0) Then AngebotNm% = LocalAngebotRec(i%).mp2
                AngebotAep# = AepKalk#
            End If
        Else
            If (LocalAngebotRec(i%).ghBest = AngebotLief%) And (LocalAngebotRec(i%).LaufNr < 1000) Then
                If (LocalAngebotRec(i%).bm <= AngebotBm%) And (LocalAngebotRec(i%).bm > MaxBm) Then
                    MaxBm = LocalAngebotRec(i%).bm
                    AngebotInd% = (i% + 1) * (-1)
'                    AngebotLief% = LocalAngebotRec(i%).ghBest
'                    AngebotBm% = LocalAngebotRec(i%).bm
'                    AngebotNm% = 0
'                    If LocalAngebotRec(i%).st = "M" Then AngebotNm% = LocalAngebotRec(i%).mp
'                    AngebotAep# = AepKalk#
                End If
            End If
        End If
    End If
Next i%

If (AngebotAuswahl%) Then
    'sortieren nach AEPkalk
    For i% = 0 To AnzLocalAngebote% - 2
      aep1# = LocalAngebotRec(i%).AepKalk
      GhPrio1% = GetGhPriorität((LocalAngebotRec(i%).gh))
      For j% = i% + 1 To AnzLocalAngebote% - 1
        aep2# = LocalAngebotRec(j%).AepKalk
        GhPrio2% = GetGhPriorität((LocalAngebotRec(j%).gh))
        If aep1# <= 0# Or aep1# > aep2# Or ((aep1# = aep2#) And (GhPrio1% > GhPrio2%)) Then
          aep1# = aep2#
          iLocalAngebotRec = LocalAngebotRec(i%)
          LocalAngebotRec(i%) = LocalAngebotRec(j%)
          LocalAngebotRec(j%) = iLocalAngebotRec
        End If
      Next j%
    Next i%
    
    iCol% = 1
    For i% = 0 To AnzLocalAngebote% - 1
        If (LocalAngebotRec(i%).Saisonal = 0) Then
            iCol% = i% + 1
            Exit For
        End If
    Next i%
    
    If (AngebotDirektEingabe%) Then
        For i% = 0 To AnzLocalAngebote% - 1
            If (LocalAngebotRec(i%).LaufNr = 2000) Then
                iCol% = i% + 1
            End If
        Next i%
    ElseIf (AngebotInd% > 0) Then
        For i% = 0 To AnzLocalAngebote% - 1
            If (LocalAngebotRec(i%).LaufNr = AngebotInd%) And (LocalAngebotRec(i%).IstManuell = AngebotManuell) Then
                iCol% = i% + 1
            End If
        Next i%
    End If

'    If (AngebotInd% >= 100) Then
    If (AngebotTemporaer) Then
'        j% = AngebotInd% - 100
        j% = iCol% - 1
        iLocalAngebotRec.gh = LocalAngebotRec(j%).gh
        iLocalAngebotRec.st = LocalAngebotRec(j%).st
        iLocalAngebotRec.bm = AngebotBm%
        If (iLocalAngebotRec.st = "P") Then
            iLocalAngebotRec.mp = LocalAngebotRec(j%).mpOrg
        Else
            iLocalAngebotRec.mp = AngebotNm%
        End If
        iLocalAngebotRec.bmOrg = LocalAngebotRec(j%).bmOrg
        iLocalAngebotRec.mpOrg = LocalAngebotRec(j%).mpOrg
        iLocalAngebotRec.ghBest = LocalAngebotRec(j%).ghBest
        iLocalAngebotRec.IstMvda = LocalAngebotRec(j%).IstMvda
        iLocalAngebotRec.Saisonal = 0
        
        For i% = AnzLocalAngebote% To 1 Step -1
            LocalAngebotRec(i%) = LocalAngebotRec(i% - 1)
        Next i%
        LocalAngebotRec(0) = iLocalAngebotRec
        AnzLocalAngebote% = AnzLocalAngebote% + 1
        iCol% = 1
    End If
    
    For i% = 0 To AnzLocalAngebote% - 1
        AepKalk# = AngebotAuswerten#(i%)
        Call AngebotZeigen(i%, AepKalk#, frmAngebote.flxAngebote)
    Next i%

    With frmAngebote.flxAngebote
        If (AnzLocalAngebote% = 0) Then
            .Cols = 2
        Else
            .Cols = AnzLocalAngebote% + 1
        End If
        
        .col = iCol%
        If (AngebotInd% > 0) Then
            .row = 0
            If (iNewLine) Then
                .CellBackColor = RGB(201, 123, 58)
            Else
                .CellBackColor = vbWhite
            End If
        End If
        .row = 1
        .RowSel = .Rows - 1
    End With

    With frmAngebote
        .lblAngeboteWert(0) = Format(BMopt!, "0.0")

'        If (Ast1.pzn <> String$(7, 0)) Then
'        If (PznLagernd%) Then
'            AepOrg2# = Ast1.aep
'        Else
'            AepOrg2# = 0#
'        End If
        .lblAngeboteWert(1) = Format(AepOrg2#, "0.00")
        
        .lblAngeboteWert(2) = Format(taxAEP#, "0.00")
    End With

End If

Call clsError.DefErrPop
End Sub

Sub AngebotZeigen(angebot%, AepKalk#, iFlex As MSFlexGrid, Optional AnzeigeEdit% = False, Optional iPicInfo As PictureBox, Optional iIndex%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotZeigen")
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
Dim gh%, GesMenge%, AnzeigeCol%, ind%
Dim AepAng#, AepAngOrg#
Dim lKurz$, h$, h2$
        
If (AnzeigeEdit%) Then
    AnzeigeCol% = 0
Else
    AnzeigeCol% = angebot% + 1
End If

lKurz$ = ""
gh% = LocalAngebotRec(angebot%).ghBest  'gh
If (LieferantenDBok) Then
    h = ""
    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(gh)
    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
    If (LieferantenRec.RecordCount <> 0) Then
        h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
    End If
    LieferantenRec.Close
    
    If (h$ = String$(Len(h$), 0)) Then h$ = ""
    If (h$ = "") Then
        h$ = "(" + Str$(gh%) + ")"
    End If
    lKurz$ = h$
'    If (LocalAngebotRec(angebot%).ghBest <> gh%) Then lKurz$ = lKurz$ + " (Mvda)"
    If (LocalAngebotRec(angebot%).IstMvda) Then lKurz$ = lKurz$ + " (Pool)"
    If (LocalAngebotRec(angebot%).LaufNr = 2000) Then
        lKurz$ = lKurz$ + " (direkt)"
    ElseIf (LocalAngebotRec(angebot%).LaufNr > 1000) Then
        lKurz$ = lKurz$ + " (allg.)"
    End If
    If (AngebotTemporaer And (AnzeigeCol% = 1)) Then lKurz$ = h$ + " (temp)"
ElseIf (gh% > 0) And (gh% <= Lif1.AnzRec) Then
    Lif1.GetRecord (gh% + 1)
    h$ = Trim$(Lif1.kurz)
    
    If (h$ = String$(Len(h$), 0)) Then h$ = ""
    If (h$ = "") Then
        h$ = "(" + Str$(gh%) + ")"
    End If
    lKurz$ = h$
'    If (LocalAngebotRec(angebot%).ghBest <> gh%) Then lKurz$ = lKurz$ + " (Mvda)"
    If (LocalAngebotRec(angebot%).IstMvda) Then lKurz$ = lKurz$ + " (Pool)"
    If (LocalAngebotRec(angebot%).LaufNr = 2000) Then
        lKurz$ = lKurz$ + " (direkt)"
    ElseIf (LocalAngebotRec(angebot%).LaufNr > 1000) Then
        lKurz$ = lKurz$ + " (allg.)"
    End If
    If (AngebotTemporaer And (AnzeigeCol% = 1)) Then lKurz$ = h$ + " (temp)"
End If

AepAng# = LocalAngebotRec(angebot%).AepAng
AepAngOrg# = LocalAngebotRec(angebot%).AepAngOrg

With iFlex
    If (AnzeigeCol% >= .Cols) Then .Cols = AnzeigeCol% + 1
    
    GesMenge% = bm% + nr%
    .TextMatrix(1, AnzeigeCol%) = lKurz$
    
    
    If (AnzeigeEdit%) Then
        h$ = "  +  999  Stk"
    Else
        h$ = Str$(OrgBm%)
        If (LocalAngebotRec(angebot%).st <> "P") And (LocalAngebotRec(angebot%).st <> "S") Then h$ = h$ + "+" + Mid$(Str$(OrgMp&), 2)
        h$ = h$ + " Stk"
    End If
    .TextMatrix(2, AnzeigeCol%) = h$
    
    h$ = Str$(bm%)
    If (nr% > 0) Then h$ = h$ + "+" + Mid$(Str$(nr%), 2)
    h$ = h$ + " Stk"
    .TextMatrix(3, AnzeigeCol%) = h$

    .TextMatrix(4, AnzeigeCol%) = Format(AEPorg#, "0.00")
    
    h$ = ""
    If (AEPorg# <> AepAngOrg#) Then h$ = Format(AepAngOrg# - AEPorg#, "0.00")
    If (h$ <> "") And (AEPorg# > 0) Then
        h$ = h$ + " (" + Format(100 - (LocalAngebotRec(angebot%).AepAngOrg / AEPorg#) * 100#, "0") + "%)"
    End If
    .TextMatrix(5, AnzeigeCol%) = h$
    
    h$ = ""
    If (AEPorg# <> AepAngOrg#) Then h$ = Format(AepAngOrg#, "0.00")
    .TextMatrix(6, AnzeigeCol%) = h$
    
    
    h$ = ""
    If (AepAngOrg# <> AepAng#) Then h$ = Format(AepAng# - AepAngOrg#, "0.00")
    .TextMatrix(7, AnzeigeCol%) = h$
    
    h$ = ""
    If (AepAngOrg# <> AepAng#) Then h$ = Format(AepAng#, "0.00")
    .TextMatrix(8, AnzeigeCol%) = h$
    
    h$ = ""
    If (nr% > 0) Then h$ = Str$(nr%) + Format(NrWert# / GesMenge%, "   (-0.00)")
    .TextMatrix(9, AnzeigeCol%) = h$
    
    .TextMatrix(10, AnzeigeCol%) = Format(LaKo# / GesMenge%, "0.00")
    .TextMatrix(11, AnzeigeCol%) = Format(PeKo# / GesMenge%, "0.00")
    .TextMatrix(12, AnzeigeCol%) = Format(StaffelAb# / GesMenge%, "0.00")
    .TextMatrix(13, AnzeigeCol%) = Format(AepKalk#, "0.00")
    .TextMatrix(14, AnzeigeCol%) = Format(gspart#, "0.00")
    .TextMatrix(15, AnzeigeCol%) = Format(prozgspart#, "0")
        
    If (AnzeigeEdit% = False) Then
        .TextMatrix(16, AnzeigeCol%) = LocalAngebotRec(angebot%).IstManuell
        .TextMatrix(17, AnzeigeCol%) = LocalAngebotRec(angebot%).LaufNr
        .TextMatrix(18, AnzeigeCol%) = LocalAngebotRec(angebot%).recno
    End If

'    If (AnzeigeEdit% = False) And (((AngebotInd% >= 100) And (angebot% = 0)) Or (angebot% = AngebotInd%)) Then
'    If (AnzeigeEdit% = False) And (((AngebotTemporaer) And (angebot% = 0)) Or (angebot% = AngebotInd%)) Then
'        .col = angebot% + 1
'        .row = 0
'        .CellBackColor = vbWhite
'    End If
End With

If (AnzeigeEdit%) And (iIndex% = 2) Then
    With iPicInfo
        .Cls
        .CurrentY = 30
        
        h$ = "ZeilRab." + vbCr + "FaktRab." + vbCr + "-" + vbCr + "GesRab."
'        h$ = "Zeilenrab." + vbCr + "Faktrab." + vbCr + "Rabatt"
        Do
            If (h$ = "") Then Exit Do
            
            ind% = InStr(h$, vbCr)
            If (ind% > 0) Then
                h2$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
            Else
                h2$ = h$
                h$ = ""
            End If
            
            .CurrentX = 30
            If (Left$(h2$, 1) = "-") Then
                iPicInfo.Line (.CurrentX, .CurrentY)-(.Width - 60, .CurrentY)
'                iPicInfo.Print
            Else
                iPicInfo.Print h2$
            End If
        Loop
        
        .CurrentY = 30
        h$ = ""
        If (ZrWert# > 0!) Then h$ = Format(ZrWert#, "0.00")
        .CurrentX = .Width - ProjektForm.TextWidth(h$) - 30
        iPicInfo.Print h$
        h$ = ""
        If (FrWert# > 0!) Then h$ = Format(FrWert#, "0.00")
        .CurrentX = .Width - ProjektForm.TextWidth(h$) - 30
        iPicInfo.Print h$
        h$ = ""
        If ((ZrWert# + FrWert#) > 0!) Then h$ = Format((ZrWert# + FrWert#), "0.00")
        .CurrentX = .Width - ProjektForm.TextWidth(h$) - 30
        iPicInfo.Print h$
    End With
End If

Call clsError.DefErrPop
End Sub

Function AngebotAuswerten#(angebot%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotAuswerten#")
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
Dim GesMenge%, dauer%, iGh%, IstDirektLieferant%, IstMvdaLieferant%, IstMvdaAngebot%, IstHapLieferant%, HatFixRabatt%, art%
Dim AepAng#, AepKalk#, AepMvda#, ZusatzAufwand#

nr% = 0: zr! = 0: fr! = 0
bm% = LocalAngebotRec(angebot%).bm
dauer% = Para1.BestellPeriode
iGh% = LocalAngebotRec(angebot%).gh

IstDirektLieferant% = 0
IstMvdaLieferant% = 0
HatFixRabatt = 0
ZusatzAufwand# = 0
If (iGh% > 0) And (iGh% <= LifZus1.AnzRec) Then
    If (AngebotModus% <> 2) Then
        LifZus1.GetRecord (iGh% + 1)
    End If
    IstDirektLieferant% = LifZus1.IstDirektLieferant
    IstMvdaLieferant% = LifZus1.IstMvdaLieferant
    HatFixRabatt = LifZus1.HatFixRabatt
    ZusatzAufwand# = LifZus1.ZusatzAufwand
    IstHapLieferant% = LifZus1.PreisBasis
End If

If (LocalAngebotRec(angebot%).st = "P") Or (LocalAngebotRec(angebot%).st = "S") Then
    If (AEPorg# > 0.009) Then
        zr! = 100 - LocalAngebotRec(angebot%).mp / AEPorg#
    End If
    
    IstMvdaAngebot% = 0
    If (LocalAngebotRec(angebot%).LaufNr <= 1000) Then
        IstMvdaAngebot% = (LocalAngebotRec(angebot%).bmOrg = 1) And (IstMvdaLieferant%)
        If (IstMvdaAngebot%) Then
            AepMvda# = LocalAngebotRec(angebot%).mp / 100#
            IstMvdaAngebot% = (AepMvda# = 0) Or (AepMvda# = AEPorg#)
        End If
    End If
    
    If (IstMvdaAngebot%) Then
'    If (bm% = 1) And (IstMvdaLieferant%) Then
        LocalAngebotRec(angebot%).AepAngOrg = AEPorg#
        zr! = LifZus1.MvdaProzent
        AepAng# = AEPorg# * (100# - (zr! + fr!)) / 100#
        LocalAngebotRec(angebot%).ghBest = LifZus1.MvdaLief
        LocalAngebotRec(angebot%).IstMvda = 1
    Else
        If (HatFixRabatt) Then
            If (InStr(Ast1.rez, "+") <> 0) Then
                HatFixRabatt = (LifZus1.FixRabatt(0) > 0)
            Else
                HatFixRabatt = (LifZus1.FixRabatt(1) > 0)
            End If
        End If
        If (HatFixRabatt) Then
            LocalAngebotRec(angebot%).AepAngOrg = AEPorg#
            If (InStr(Ast1.rez, "+") <> 0) Then
                zr! = LifZus1.FixRabatt(0)
            Else
                zr! = LifZus1.FixRabatt(1)
            End If
            AepAng# = AEPorg# * (100# - (zr! + fr!)) / 100#
        Else
            AepAng# = LocalAngebotRec(angebot%).mp / 100#
            LocalAngebotRec(angebot%).AepAngOrg = AepAng#
        End If
    End If
'    If (LifZus1.CheckAusnahme(-1, 2, bm%, AEPorg#)) Then fr! = LifZus1.PrognoseRabatt
    art% = 2
    If (LocalAngebotRec(angebot).LaufNr >= 1000) Then art% = 0
    fr! = LifZus1.GlobalPrognoseRabatt(art%, bm%, AEPorg#)
    AepAng# = AepAng# * (100# - fr!) / 100#
    
    If (Para1.Land = "A") Then
        nr = LocalAngebotRec(angebot).mp2
    End If
Else
    nr% = LocalAngebotRec(angebot%).mp
    zr! = LocalAngebotRec(angebot%).zr
    
    If (IstDirektLieferant%) Then
        If (zr! < 0) Then
            fr! = LifZus1.FakturenRabatt
            If (LifZus1.TempFakturenRabatt > 0) Then fr! = LifZus1.TempFakturenRabatt
            zr! = zr! * (-1)
            If (zr! = 123.45!) Then zr! = 0
            If (zr! = 200!) Then zr! = 0
        End If
        
        If (LocalAngebotRec(angebot).LaufNr >= 1000) And (IstHapLieferant%) Then
            LocalAngebotRec(angebot%).AepAngOrg = TaxeHAP#
        End If
        
        If (LocalAngebotRec(angebot%).AepAngOrg > 0) Then
            If (AEPorg# > 0) Then
                zr! = 100 - (LocalAngebotRec(angebot%).AepAngOrg / AEPorg#) * 100#
            Else
                zr! = 0
            End If
'            If (IstHapLieferant%) Then
'                LocalAngebotRec(angebot%).zr = zr!
'            End If
        Else
            LocalAngebotRec(angebot%).AepAngOrg = AEPorg# * (100# - Abs(zr!)) / 100#
        End If
        
        If (LifZus1.BevorratungsZeitraum > 0) Then dauer% = LifZus1.BevorratungsZeitraum
        If (LifZus1.TempBevorratungsZeitraum > 0) Then dauer% = LifZus1.TempBevorratungsZeitraum
        
'        If (LocalAngebotRec(angebot).LaufNr >= 1000) And (IstHapLieferant%) Then
            AepAng# = LocalAngebotRec(angebot%).AepAngOrg * (100# - fr!) / 100#
'        Else
'            If (AEPorg# > 0) Then
'                AepAng# = AEPorg# * (100# - (zr! + fr!)) / 100#
'            Else
'                AepAng# = LocalAngebotRec(angebot%).AepAngOrg * (100# - (zr! + fr!)) / 100#
'            End If
'        End If
    Else
        LocalAngebotRec(angebot%).AepAngOrg = AEPorg#
        
        art% = 2
        If (LocalAngebotRec(angebot).LaufNr >= 1000) Then art% = 0
'        If (LifZus1.CheckAusnahme(-1, art%, bm%, AEPorg#)) Then fr! = LifZus1.PrognoseRabatt
        fr! = LifZus1.GlobalPrognoseRabatt(art%, bm%, AEPorg#)
    
        If (AEPorg# > 0) Then
            AepAng# = AEPorg# * (100# - (zr! + fr!)) / 100#
        Else
            AepAng# = LocalAngebotRec(angebot%).AepAngOrg * (100# - (zr! + fr!)) / 100#
        End If
    End If
End If
LocalAngebotRec(angebot%).AepAng = AepAng#
OrgBm% = LocalAngebotRec(angebot%).bmOrg
OrgMp& = LocalAngebotRec(angebot%).mpOrg
Call AngebotRechnen(bm%, nr%, zr!, fr!, AepAng#, dauer%, iGh%, IstDirektLieferant%, ZusatzAufwand#)
GesMenge% = bm% + nr%
If (bm% = 1) Then
    AepKalk# = AepAng#
    LaKo# = 0
    PeKo# = 0
Else
    AepKalk = AepAng + Manko# + (StaffelAb# + LaKo# + PeKo# - NrWert#) / GesMenge%
End If
prozgspart# = 0#
If AEPorg# > 0.009 Then
  gspart# = (AEPorg# - AepKalk#) * GesMenge%
  prozgspart# = (AEPorg# - AepKalk#) / AEPorg# * 100#
Else
  gspart# = 0
  prozgspart# = 0
End If

LocalAngebotRec(angebot%).AepKalk = AepKalk#
AngebotAuswerten# = AepKalk#

Call clsError.DefErrPop
End Function

Static Sub AngebotRechnen(bm%, nr%, zr!, fr!, AepAng#, dauer%, gh%, IstDirektLief%, ZusatzAufwand#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotRechnen")
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
Dim Wert#, lx#, LaZi#
Dim zuwenig!, zuviel!, WieLange!, RABMIT!, proz!
Dim i%, GesMenge%, RabSatz%

LaKo# = 0#
PeKo# = 0#
LaZi# = 0#

If (bm% <> 0) Or (nr% <> 0) Then

'  If (BMopt! <= 0) Then BMopt! = BMast%

  GesMenge% = bm% + nr%

  '* Bestellwert errechnen
  Wert# = CDbl(bm%) * AepAng#

  '* Naturalrabattwert (mit Faktor)
  NrWert# = CDbl(nr%) * AEPorg# * Para1.FakNatu

  '* Zeilen-Barrabattwert (mit Faktor)
  ZrWert# = (zr! / 100#) * AEPorg#
  FrWert# = (fr! / 100#) * AEPorg#
  BarRabWert# = ((CDbl(zr!) + fr!) / 100#) * CDbl(bm%) * AEPorg# * Para1.FakBar
'  BarRabWert# = ((CDbl(zr!) + fr!) / 100#) * GesMenge% * AEPorg# * Para1.FakBar

'  If (IstDirektLief% = False) Or (BMoptOrg! > 0) Then
    '* zuviel bestellt
    If GesMenge% > BMopt! Then
      zuviel! = CSng(GesMenge%) - BMopt!
      WieLange! = (zuviel! / BMopt!) * Para1.BestellPeriode 'dauer
      LaKo# = WieLange! / 360# * (Para1.Lagerkosten) / 100# * AEPorg# * zuviel! 'AepAng#
      PeKo# = WieLange! / 360# * (Para1.PersonalKosten) / 100# * AEPorg# * zuviel! 'AepAng#
      LaZi# = LaKo# + PeKo#
    End If

    '* zuwenig bestellt
    If GesMenge% < BMopt! Then
      zuwenig! = BMopt! - CSng(GesMenge%)
      WieLange! = (zuwenig! / BMopt!) * Para1.BestellPeriode    'dauer
      lx# = WieLange! / 360# * (Para1.PersonalKosten) / 100# * AEPorg# * zuwenig! 'AepAng#
      LaZi# = LaZi# + lx#
      PeKo# = PeKo# + lx#
    End If
'  End If
  
  PeKo# = PeKo# + (ZusatzAufwand# / 100#) * AEPorg# ' * AepAng#
  
  Manko# = (100# - Para1.Tfaktor) * AEPorg# / 100#  'AepAng#
  
  StaffelAb# = 0#
'    If (IstDirektLief%) Then
'        RabSatz% = 0
'        For i% = 1 To Para1.StaffelMax - 2
'            If (GesMenge% >= Para1.StaffelBM(i%)) And (GesMenge% <= Para1.StaffelBM(i% + 1)) Then
'                RabSatz% = i%
'                Exit For
'            End If
'        Next i%
'        If (RabSatz% > 0) Then
'            'rabatt zwischen RabSatz% und RabSatz%+1
'            proz! = (GesMenge% - Para1.StaffelBM(RabSatz%)) / (Para1.StaffelBM(RabSatz% + 1) - Para1.StaffelBM(RabSatz%))
'            RABMIT! = Para1.StaffelRabatt(RabSatz%) + (Para1.StaffelRabatt(RabSatz% + 1) - Para1.StaffelRabatt(RabSatz%)) * proz! 'linear
'        Else
'            RABMIT! = Para1.StaffelRabatt(Para1.StaffelMax)
'        End If
'
'        StaffelAb# = GesMenge% * AEPorg# * RABMIT! / 100# * Para1.FakBar
'    ElseIf (BMast% > 0) Then
'        If (bm% + nr%) > BMast% Then
'          StaffelAb# = ((bm% + nr%) - BMast%) / BMast% * Para1.PersonalKosten / 100# * AepAng#
'        End If
'    End If
End If
    
Call clsError.DefErrPop
End Sub

Sub DirektBewertung(modus%, Optional ind% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DirektBewertung")
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
Dim i%, AngebotDa%, GesMenge%, gh%
Dim TaxeAep#, Bewertung#, vt#
Dim asatz&
Dim pzn$, h$, SQLStr$, lKurz$
Dim iFlex As MSFlexGrid

If (modus% = 0) Then
    GesBm% = 0
    GesNr% = 0
    GesAepOrg# = 0
    GesAepScreen# = 0
    GesNrWert# = 0
    GesBarRabWert# = 0
    GesBarRabScreen# = 0
    GesLaKo# = 0
    GesPeKo# = 0
    GesStaffelAb# = 0
    GesProzGspart# = 0
ElseIf (modus% = 1) Then
    BMopt! = AngebotBMopt!
    BMast% = AngebotBMast%
    
    pzn$ = AngebotPzn$
    
    TaxeAep# = AngebotTaxeAep#
    WgAst% = AngebotWgAst%
    
    AngebotModus% = 2
    Call AngebotSelect(pzn$, TaxeAep#, AngebotPznLagernd%)
    
    GesMenge% = bm% + nr%
    
    GesBm% = GesBm% + bm%
    GesNr% = GesNr% + nr%
    GesAepScreen# = GesAepScreen# + (AEPorg# * bm%)
    GesAepOrg# = GesAepOrg# + (AEPorg# * GesMenge%)
    GesNrWert# = GesNrWert# + NrWert#
    GesBarRabScreen# = GesBarRabScreen# + BarRabWert#
    GesBarRabWert# = GesBarRabWert# + BarRabWert# + (ZrWert# + FrWert#) * CDbl(nr%) * Para1.FakBar
    GesLaKo# = GesLaKo# + LaKo#
    GesPeKo# = GesPeKo# + PeKo#
    GesStaffelAb# = GesStaffelAb# + StaffelAb#
Else
    GesAepKalk# = GesAepOrg# - GesBarRabWert# - GesNrWert#
    GesAepKalk# = GesAepKalk# + GesStaffelAb# + GesLaKo# + GesPeKo#
    
    GesGspart# = GesAepOrg# - GesAepKalk#
    If (GesAepOrg# > 0#) Then
        GesProzGspart# = GesGspart# / GesAepOrg# * 100#
    
        Bewertung# = GesNrWert# / (GesAepScreen# + GesNrWert#)
        Bewertung# = Bewertung# + (GesBarRabScreen# - GesLaKo# - GesPeKo# - GesStaffelAb#) / GesAepScreen#

        If (LifZus1.ValutaStellung > 0) Then vt# = LifZus1.ValutaStellung
        If (LifZus1.TempValutaStellung > 0) Then vt# = LifZus1.TempValutaStellung
        Bewertung# = Bewertung# + (vt# / 360#) * Para1.Lagerkosten / 100#
        
        Bewertung# = Bewertung# * 100#
    End If
    Bewertung# = Bewertung# + 100#
'    Bewertung# = Bewertung# - Para1.TransDataRabatt * Para1.FakBar
    
    BewertungOk% = 0 ' True
    If (GesAepScreen# < LifZus1.DirektMindestWert) Then BewertungOk% = BewertungOk% Or &H1  ' False
    If (Bewertung# < LifZus1.DirektMindestBewertung) Then BewertungOk% = BewertungOk% Or &H2  ' False
    
    GesBewertung# = Bewertung#
 
    If (ind%) Then
       
        AngebotModus% = 2
        
        If (ind% = -1) Then
            Load frmAngebote
        
            With frmAngebote
                .Caption = "Bewertung Direktbezug"
                .Left = (Screen.Width - .Width) / 2
                If (AngebotY% < 0) Then AngebotY% = (Screen.Height - .Height) / 2
                .Top = AngebotY%
            End With
            
            Set iFlex = frmAngebote.flxAngebote
        Else
            Set iFlex = ProjektForm.flxeinzelbewertung
        End If
    
        With iFlex
            
            lKurz$ = ""
            gh% = AngebotActLief%
            If (LieferantenDBok) Then
                h = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(gh)
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                End If
                LieferantenRec.Close
                If (h$ = String$(Len(h$), 0)) Then h$ = ""
                If (h$ = "") Then
                    h$ = "(" + Str$(gh%) + ")"
                End If
                lKurz$ = h$
            ElseIf (gh% > 0) And (gh% <= Lif1.AnzRec) Then
                Lif1.GetRecord (gh% + 1)
                h$ = Trim$(Lif1.kurz)
                If (h$ = String$(Len(h$), 0)) Then h$ = ""
                If (h$ = "") Then
                    h$ = "(" + Str$(gh%) + ")"
                End If
                lKurz$ = h$
            End If
            .TextMatrix(1, 1) = lKurz$
    
            h$ = Str$(GesBm%)
            If (GesNr% > 0) Then h$ = h$ + "+" + Mid$(Str$(GesNr%), 2)
            h$ = h$ + " Stk"
            .TextMatrix(3, 1) = h$
        
            .TextMatrix(4, 1) = Format(GesAepScreen#, "0.00")
            
            h$ = ""
            If (GesBarRabScreen# > 0) Then h$ = Format(-GesBarRabScreen#, "0.00")
            .TextMatrix(5, 1) = h$
            
            h$ = ""
            If (GesBarRabScreen# > 0) Then h$ = Format(GesAepScreen# - GesBarRabScreen#, "0.00")
            .TextMatrix(6, 1) = h$
            
            h$ = ""
            If (GesNr% > 0) Then h$ = Str$(GesNr%) + Format(GesNrWert#, "   (-0.00)")
            .TextMatrix(7, 1) = h$
            
            .TextMatrix(8, 1) = Format(GesLaKo#, "0.00")
            .TextMatrix(9, 1) = Format(GesPeKo#, "0.00")
            .TextMatrix(10, 1) = Format(GesStaffelAb#, "0.00")
            
            
    '        GesAepKalk# = GesAepKalk# - GesAepWert# + GesAepOrg# + GesBarRabWert# - GesBarRabWertOrg#
            GesAepKalk# = GesAepKalk# - GesAepOrg# + GesAepScreen#
            .TextMatrix(11, 1) = Format(GesAepKalk#, "0.00")
            
            
            .TextMatrix(12, 1) = Format(GesGspart#, "0.00")
            .TextMatrix(13, 1) = Format(GesProzGspart#, "0") + Format(GesBewertung#, "   (0.00)")
            
            .col = 1
            .row = 1
            .RowSel = .Rows - 1
        
        End With
        
        If (ind% = -1) Then frmAngebote.Show 1
    End If

End If

Call clsError.DefErrPop
End Sub

Function GetGhPriorität%(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetGhPriorität%")
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
            
ret% = InStr(AngebotGhPriorität$, "," + Format(iLief%, "000") + ",")
If (ret% = 0) Then
    ret% = 999
End If
GetGhPriorität = ret%

Call clsError.DefErrPop
End Function

