VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDialoge"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function MatchCode%(mTyp%, pzn$, txt$)
Dim ret%

ret% = False

If (mTyp% = 0) Then
    Set Match1 = New clsMatchArtikel
    MatchTyp% = MATCH_ARTIKEL
Else
    Set Match1 = New clsMatchLieferanten
    MatchTyp% = MATCH_LIEFERANTEN
End If
MatchcodePzn$ = "0000000"
MatchcodeTxt$ = Trim(txt$)
Call Match1.InitMatch
frmMatchcode.Show 1
pzn$ = MatchcodePzn$
txt$ = MatchcodeTxt$
If (pzn$ <> "0000000") Then
    ret% = True
End If

Set Match1 = Nothing
MatchTyp% = -1

MatchCode% = ret%

End Function

Sub AnzeigeFenster(typ$, key As Variant, txt$)
Dim lief%
Dim pzn$

AnzeigeFensterTyp$ = typ$
Load frmBstatus
If (typ$ = TEXT_BESTELLSTATUS) Then
    pzn$ = key
    Call BstatusBefuellen(pzn$)
ElseIf (typ$ = TEXT_NACHBEARBEITUNG) Then
    lief% = key
    Call NachbearbeitungBefuellen(lief%)
ElseIf (typ$ = TEXT_ABSAGEN) Then
    lief% = key
    Call AbsagenBefuellen(lief%)
End If

With frmBstatus
    .Caption = typ$ + " - " + txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmBstatus.Show 1

End Sub

Sub BestellStatus(pzn$, txt$)

Load frmBstatus
Call BstatusBefuellen(pzn$)
'
With frmBstatus
    .Caption = "Bestellstatus - " + txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmBstatus.Show 1

End Sub

Sub BstatusBefuellen(pzn$)
Dim i%, erst%, Max%, lief%, iZeit%, alm1%, arm1%, m%
Dim FabsRecno&
Dim h$, h2$, nbtext$

With frmBstatus.flxbstatus
    .row = 0
    
    erst% = True
    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Ass1.GetRecord (FabsRecno& + 1)
        lief% = Ass1.llief(4)
        h$ = Format(lief%, "###    ")
        If (lief% > 0) And (lief% < 200) Then
            Lif1.GetRecord (lief% + 1)
            h2$ = Lif1.kurz
            Call OemToChar(h2$, h2$)
            h$ = h$ + h2$
        End If
        If (erst% = True) Then
            .Rows = .Rows + 1
            .row = .Rows - 1
            .col = 0
            .CellFontBold = True
            .TextMatrix(.row, 0) = "Letzte Lieferung"
            erst% = False
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
    
    erst% = True
    Bek1.GetRecord (1)
    Max% = Bek1.erstmax
    For i% = 1 To Max%
        Bek1.GetRecord (i% + 1)
        If (pzn$ = Bek1.pzn) Then
            lief% = Bek1.lief
            h$ = Format(lief%, "###    ")
            If (lief% > 0) And (lief% < 200) Then
                Lif1.GetRecord (lief% + 1)
                h2$ = Lif1.kurz
                Call OemToChar(h2$, h2$)
                h$ = h$ + h2$
            End If
            If (erst% = True) Then
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = True
                .TextMatrix(.row, 0) = "Bestellung"
                erst% = False
            End If
            .Rows = .Rows + 1
            .row = .Rows - 1
            .col = 0
            .CellFontBold = False
            .TextMatrix(.row, 3) = h$
            .TextMatrix(.row, 4) = Abs(Bek1.bm) + Abs(Bek1.nm)
        End If
    Next i%
    
    erst% = True
    Call Wu1.GetRecord(1)
    Max% = Wu1.erstmax
    For i% = 1 To Max%
        Call Wu1.GetRecord(i% + 1)
        If (pzn$ = Left$(Wu1.at, 7)) Then
            lief% = Asc(Wu1.li)
            h$ = Format(lief%, "###    ")
            If (lief% > 0) And (lief% < 200) Then
                Lif1.GetRecord (lief% + 1)
                h2$ = Lif1.kurz
                Call OemToChar(h2$, h2$)
                h$ = h$ + h2$
            End If
            If (erst% = True) Then
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = True
                .TextMatrix(.row, 0) = "Warenübernahme"
                erst% = False
            End If
            .Rows = .Rows + 1
            .row = .Rows - 1
            .col = 0
            .CellFontBold = False
            .TextMatrix(.row, 1) = Left$(Wu1.bd, 2) + "." + Mid$(Wu1.bd, 3, 2) + "." + Mid$(Wu1.bd, 5, 2)
            iZeit% = clsOpTool.CVI(Wu1.zeit)
            .TextMatrix(.row, 2) = Format(iZeit% \ 100, "00") + ":" + Format(iZeit% Mod 100, "00")
'            .TextMatrix(.row, 2) = Format(best.zeit \ 100, "00") + ":" + Format(best.zeit Mod 100, "00")
            .TextMatrix(.row, 3) = h$
            .TextMatrix(.row, 4) = Wu1.bm + Wu1.nm
        End If
    Next i%
    
    erst% = True
    Nb1.OpenDatei
    Nb1.GetRecord (1)
    Max% = Nb1.erstmax
    
    For i% = 1 To Max%
        Nb1.GetRecord (i% + 1)
        If (pzn$ = Left$(Nb1.at, 7)) Then
            alm1% = Nb1.alm: arm1% = Nb1.arm: nbtext$ = ""
            If (alm1% = 0) And (arm1% = 0) Then
                alm1% = Nb1.bm + Nb1.nm - Nb1.lm
                If (alm1% <> 0) Then nbtext$ = "unbestätigt"
            ElseIf (alm1% > 0) Then
                nbtext$ = "Nachlieferung"
            ElseIf (alm1% < 0) Then
                nbtext$ = "Retourware"
            End If
            If (nbtext$ <> "") Then
                lief% = Asc(Nb1.li)
                h$ = Format(lief%, "###    ")
                If (lief% > 0) And (lief% < 200) Then
                    Lif1.GetRecord (lief% + 1)
                    h2$ = Lif1.kurz
                    Call OemToChar(h2$, h2$)
                    h$ = h$ + h2$
                End If
                If (erst% = True) Then
                    .Rows = .Rows + 1
                    .row = .Rows - 1
                    .col = 0
                    .CellFontBold = True
                    .TextMatrix(.row, 0) = "Nachbearbeitung"
                    erst% = False
                End If
                .Rows = .Rows + 1
                .row = .Rows - 1
                .col = 0
                .CellFontBold = False
                .TextMatrix(.row, 0) = nbtext$
                .TextMatrix(.row, 1) = Left$(Nb1.bd, 2) + "." + Mid$(Nb1.bd, 3, 2) + "." + Mid$(Nb1.bd, 5, 2)
                .TextMatrix(.row, 2) = ""
                .TextMatrix(.row, 3) = h$
                .TextMatrix(.row, 4) = alm1%
            End If
        End If
    Next i%
    
    Nb1.CloseDatei
    
    
    erst% = True
    
'neu ab 1.0.9
    FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
    Do
        If (FabsErrf%) Then Exit Do

        Besorgt1.GetRecord (FabsRecno& + 1)

        If (pzn$ <> Besorgt1.pzn) Then Exit Do

        m% = Besorgt1.lm
        lief% = Besorgt1.lief
        h$ = Format(lief%, "###    ")
        If (lief% > 0) And (lief% < 200) Then
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
        If (erst% = True) Then
            .Rows = .Rows + 1
            .row = .Rows - 1
            .col = 0
            .CellFontBold = True
            .TextMatrix(.row, 0) = "Merkzettel"
            erst% = False
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
'Ende neu ab 1.0.9
    
    If (.Rows = 1) Then
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 0) = "keine Daten gefunden !"
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

End Sub

Sub NachbearbeitungBefuellen(lief%)
Dim i%, erst%, Max%, iZeit%, alm1%, arm1%, m%
Dim FabsRecno&
Dim h$, h2$, nbtext$

With frmBstatus.flxbstatus
    .row = 0
    
    erst% = True
    Nb1.OpenDatei
    Nb1.GetRecord (1)
    Max% = Nb1.erstmax
   
    For i% = 1 To Max%
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

End Sub

Sub AbsagenBefuellen(lief%)
Dim i%, erst%, Max%, iZeit%, alm1%, arm1%, m%
Dim FabsRecno&
Dim h$, h2$, nbtext$

With frmBstatus.flxbstatus
    .row = 0
    
    erst% = True
    Absagen1.OpenDatei
    Absagen1.GetRecord (1)
    Max% = Absagen1.erstmax
   
    For i% = 1 To Max%
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
    
    If (.Rows = 1) Then
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 2) = "keine Daten gefunden !"
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

End Sub

Function TesteBstatus%(pzn$)
Dim i%, Max%
Dim FabsRecno&

TesteBstatus% = True
    
FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Exit Function
End If
    
Bek1.GetRecord (1)
Max% = Bek1.erstmax
For i% = 1 To Max%
    Bek1.GetRecord (i% + 1)
    If (pzn$ = Bek1.pzn) Then
        Exit Function
    End If
Next i%
    
Call Wu1.GetRecord(1)
Max% = Wu1.erstmax
For i% = 1 To Max%
    Call Wu1.GetRecord(i% + 1)
    If (pzn$ = Left$(Wu1.at, 7)) Then
        Exit Function
    End If
Next i%
    
Nb1.OpenDatei
Nb1.GetRecord (1)
Max% = Nb1.erstmax
For i% = 1 To Max%
    Nb1.GetRecord (i% + 1)
    If (pzn$ = Left$(Nb1.at, 7)) Then
        Nb1.CloseDatei
        Exit Function
    End If
Next i%

FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Exit Function
End If
    
TesteBstatus% = False

End Function

Sub ZusatzFenster(typ$, key As Variant, txt$)
Dim lief%
Dim pzn$

ZusatzFensterTyp$ = typ$
Load frmZusatz
If (typ$ = ZUSATZ_ARTIKEL) Then
    pzn$ = key
    ZusatzPzn$ = pzn$
    Call ZusatzArtikelBefuellen(pzn$)
ElseIf (typ$ = ZUSATZ_LIEFERANTEN) Then
    lief% = key
    Call ZusatzLieferantenBefuellen(lief%)
End If

With frmZusatz
    .Caption = txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmZusatz.Show 1

End Sub

Sub ZusatzArtikelBefuellen(pzn$)
Dim i%
Dim FabsRecno&
Dim txt$

With frmZusatz.flxZusatz
    FabsErrf% = Zus1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Zus1.GetRecord (FabsRecno& + 1)
        For i% = 0 To 4
            txt$ = Trim$(Zus1.zusatz(i%))
            Call OemToChar(txt$, txt$)
            .TextMatrix(i% + 1, 0) = txt$
            frmZusatz.txtZusatz(i%).text = txt$
        Next i%
    End If

    .row = 1
End With

End Sub

Sub ZusatzLieferantenBefuellen(lief%)
Dim i%
Dim FabsRecno&
Dim txt$, key$

With frmZusatz.flxZusatz
    key$ = Format(lief%, "00000") + "001"
    FabsErrf% = lZus1.IndexSearch(0, key$, FabsRecno&)
    For i% = 0 To 99
        If (FabsErrf% <> 0) Then
            Exit For
        End If
        
        lZus1.GetRecord (FabsRecno& + 1)
        
        If (lZus1.LiefNr <> lief%) Then
            Exit For
        End If
        
        txt$ = Trim$(lZus1.text)
        Call OemToChar(txt$, txt$)
        .TextMatrix(i% + 1, 0) = txt$
'        frmZusatz.txtZusatz(i%).text = txt$
    
        key$ = clsfabs.LastKey(8)
        FabsErrf% = lZus1.IndexNext(0, FabsRecno&, RTrim$(key$), FabsRecno&)
    Next i%

    .row = 1
End With

End Sub

Sub ZusatzSpeichern()
Dim i%
Dim FabsRecno&
Dim txt$

FabsErrf% = Zus1.IndexSearch(0, ZusatzPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Zus1.GetRecord (FabsRecno& + 1)
Else
    FabsErrf% = Zus1.IndexInsert(0, ZusatzPzn$, FabsRecno&)
End If
If (FabsErrf% = 0) Then
    Zus1.pzn = ZusatzPzn$
    For i% = 0 To 4
        txt$ = Trim$(frmZusatz.txtZusatz(i%).text)
        Call CharToOem(txt$, txt$)
        Zus1.zusatz(i%) = txt$
    Next i%
    Zus1.PutRecord (FabsRecno& + 1)
End If

End Sub

Sub BesorgerInfo(Nummer%, pzn$, txt$)
Dim erg%

Set Kiste1 = New clsKiste
erg% = AbholerNummer%(Nummer%, pzn$)
If (erg%) Then
    BesorgerNummer% = Nummer%
    BesorgerPzn$ = pzn$
    Load frmAbholer
    With frmAbholer
        .Caption = "Abholnummer " + Str$(Nummer%) + " - " + txt$
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        .Show 1
    End With
End If
Set Kiste1 = Nothing

End Sub

Sub BesorgerBefuellen()
Dim erg%, ret%, PersCode%, KundenNr%
Dim h$, tx$, RezTxt$

Call SucheBesorgerInVk(BesorgerNummer%, PersCode%, KundenNr%)

With frmAbholer
    tx$ = ""
    If (PersCode% > 0) Then
        tx$ = Para1.Personal(PersCode%)
    End If
    .lblAbholerWert(0).Caption = tx$
    
    tx$ = Kiste1.VonWo \ 2
    If (Kiste1.VonWo Mod 2) Then
        tx$ = tx$ + " Rechts"
    Else
        tx$ = tx$ + " Links"
    End If
    .lblAbholerWert(1).Caption = tx$
    
    tx$ = clsOpTool.CVDatum(Left$(Kiste1.VonWann, 2))
    .lblAbholerWert(2).Caption = Mid$(tx$, 7, 2) + "." + Mid$(tx$, 5, 2) + "." + Left$(tx$, 4)
    
    .lblAbholerWert(3).Caption = Format(Asc(Mid$(Kiste1.VonWann, 3, 1)), "00") + ":" + Format(Asc(Mid$(Kiste1.VonWann, 4, 1)), "00")
    
    tx$ = ""
    If (KundenNr% > 0) Then
        tx$ = clsOpTool.HoleKundenName(KundenNr%)
    End If
    .lblAbholerWert(4).Caption = tx$
    
    tx$ = ""
    If (Kiste1.RezeptNr) > 0 Then
        RezTxt$ = clsOpTool.Bcd2ascii(Kiste1.RezeptEan, 12)
        RezTxt$ = RezTxt$ + "0": Call clsOpTool.EanPruef(RezTxt$)
        tx$ = RezTxt$
    End If
    .lblAbholerWert(5).Caption = tx$
    
    .lblAbholerWert(6).Caption = LTrim$(BesorgerAconto$)
    .lblAbholerWert(7).Caption = Str$(BesorgerMenge%)
    .lblAbholerWert(8).Caption = RTrim$(Kiste1.InfoText(3))
    .lblAbholerWert(9).Caption = RTrim$(Kiste1.InfoText(4))
End With

End Sub

Function AbholerNummer%(Nummer%, pzn$)
Dim i%, j%, satz1%, satz%, header%, Bit%, bitmaske%, frei%, kistefertig%, allesunfertig%, menge%
Dim text%, InfoText%, PznGefunden%, Belegt%, BlockNummer%
Dim KistenSatz&
Dim preis#
Dim ch$, Arttext$, aiPzn$, aconto$, txt$, wg$, mText$, mx$, h$, h1$, h2$

AbholerNummer% = False

If Abs(Nummer%) > 999 Then
    Exit Function
End If
If Nummer% < 0 Then
    Exit Function
End If

Kiste1.OpenDatei
Belegt% = Kiste1.Belegt(Nummer%)
If (Belegt%) Then
    BlockNummer% = Kiste1.PznInKiste(Nummer%, pzn$)
    If (BlockNummer% >= 0) Then
        Kiste1.GetInhalt (BlockNummer%)
        For j% = 0 To 9
            h$ = RTrim$(Kiste1.InfoText(j%))
            h1$ = Mid$(h$, 6, 2)
            h2$ = Left$(h$, 7)
            If (j% <= 5) And (h1$ = "A=" Or h1$ = "G=" Or h2$ = "ACONTO=" Or h2$ = "GEBÜHR=") Then
                BesorgerAconto$ = Mid$(h$, 8, 10)
                BesorgerMenge% = Val(Mid$(h$, 21, 4))
                If (BesorgerMenge% = 0) Then BesorgerMenge% = 1
                Exit For
            End If
        Next j%
        AbholerNummer% = True
    End If
End If

Kiste1.CloseDatei

End Function

Sub SucheBesorgerInVk(AbholNr%, PersCode%, KundenNr%)
Dim i%, j%, VERKAUF%
Dim von&, bis&, index&, VerkaufMax&, ret&
Dim h$, s$, SuchDatum$, SuchZeit$, Such$, SuchKiste$

PersCode% = 0
KundenNr% = 0

SuchDatum$ = Left$(Kiste1.VonWann, 2)
SuchZeit$ = clsOpTool.MKI(Asc(Mid$(Kiste1.VonWann, 3, 1)) * 100 + Asc(Mid$(Kiste1.VonWann, 4, 1)))
Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)

SuchKiste$ = Format(AbholNr%, "0")

For i% = 1 To 4
    Debug.Print Asc(Mid$(Such$, i%, 1));
Next i%
Debug.Print
Debug.Print
                
h$ = "verkauf.dat"
VERKAUF% = clsDat.FileOpen(h$, "R", "R", 128)

VerkaufMax& = (LOF(VERKAUF%) / 128) - 1

For j% = 0 To 1
    If (j% = 0) Then
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)
    Else
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Chr$(Asc(Mid$(SuchZeit$, 1, 1)) - 1)
    End If
    
    von& = 1&
    bis& = VerkaufMax&
    
    Do While (von& <= bis&)
        index& = (von& + bis&) \ 2
        Get #VERKAUF%, index& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        Debug.Print index&;
        For i% = 1 To 4
            Debug.Print Asc(Mid$(s$, i%, 1));
        Next i%
        Debug.Print
        If (Such$ = s$) Then
            ret& = index&
            Exit Do
        ElseIf (Such$ < s$) Then
            bis& = index& - 1
        Else
            von& = index& + 1
        End If
    Loop
    
    If (ret& >= 1) Then Exit For
Next j%

If (ret& >= 1) Then
    Do
        ret& = ret& - 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        If (Such$ <> s$) Then
            Exit Do
        End If
    Loop
    
    Do
        ret& = ret& + 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        If (Such$ <> s$) Then
            Exit Do
        End If
        If (InStr(VkRec.text, "Abhol-Nr") > 0) And (InStr(VkRec.text, SuchKiste$) > 0) Then
            KundenNr% = VkRec.knr
            PersCode% = VkRec.PersCode
            Exit Do
        End If
    Loop
End If

Close #VERKAUF%

End Sub

Sub ZeigeAngebote(pzn$, txt$, angModus%, angInd%, Optional angBm%, Optional angNm%, Optional angLief%, Optional angAep#)
Dim iAng%

AngebotPzn$ = pzn$
AngebotModus% = angModus%
AngebotInd% = angInd%
AngebotAuswahl% = True

Load frmAngebote
With frmAngebote
    .Caption = "GH-Angebote - " + txt$
    .Left = (Screen.Width - .Width) / 2
    .Top = (Screen.Height - .Height) / 2
End With
frmAngebote.Show 1

If (AngebotInd% > 0) Then
    iAng% = AngebotInd% - 1
    angBm% = LocalAngebotRec(iAng%).bm
    If (LocalAngebotRec(iAng%).st = "P") Then
        angNm% = 0
    Else
        angNm% = LocalAngebotRec(iAng%).mp
    End If
    angLief% = LocalAngebotRec(iAng%).gh
    angAep# = LocalAngebotRec(iAng%).AepKalk
    angInd% = AngebotInd%
End If
angInd% = AngebotInd%

End Sub

Function ErstAngebot%(pzn$, angInd%, angBm%, angNm%, angLief%, angAep#)
Dim ret%, ind%

ret% = False

AngebotPzn$ = pzn$
AngebotAuswahl% = False
AngebotLief% = angLief%

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
            angNm% = LocalAngebotRec(ind%).mp
        End If
        angLief% = LocalAngebotRec(ind%).gh
        angAep# = LocalAngebotRec(ind%).AepKalk
    End If
Else
    angInd% = 1
End If

ErstAngebot% = ret%

End Function

Sub ZeigeStatbild(SuchPzn$, pForm As Object)
Dim TaskId&, FabsRecno&
Dim STATBILD%, erg%, FabsErrf%
Dim dPreis#
Dim s$, pzn$, ch$, SQLStr$
Dim recTaxe As Recordset

If (SuchPzn$ = "") Then Exit Sub

FabsErrf% = Ass1.IndexSearch(0, SuchPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Ass1.GetRecord (FabsRecno& + 1)
    
    FabsErrf% = Ast1.IndexSearch(0, SuchPzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Ast1.GetRecord (FabsRecno& + 1)
    
        STATBILD% = clsDat.FileOpen("statb" + Para1.user + ".$$$", "W")
'        STATBILD% = FileOpen%("statbild.$$$", "W")
        Ast1.PrintFile (STATBILD%)
        Ass1.PrintFile (STATBILD%)
'        Put STATBILD%, , Ast1
'        Put STATBILD%, , Ass1
        
        s$ = String$(200, 0)
        Put STATBILD%, , s$
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + SuchPzn$
        Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
        If (recTaxe.EOF = False) Then
            dPreis# = recTaxe!EK / 100
            Call DxToMBFd(dPreis#)
            Put STATBILD%, , dPreis#
            dPreis# = recTaxe!VK / 100
            Call DxToMBFd(dPreis#)
            Put STATBILD%, , dPreis#
        End If
        
        Put STATBILD%, , Left$(s$, 2)
        
        Close #STATBILD%
        
        TaskId& = Shell("\user\dosrun.bat " + Para1.user + " statbild.exe", vbNormalFocus)
        Call clsOpTool.WarteAufTaskEnde(TaskId&, pForm)
'        AppActivate Me
        
    End If
End If
        
End Sub

Function WechselFenster%(zeilen$(), start%)
Dim i%, Max%

Max% = UBound(zeilen$)
ReDim WechselZeilen$(Max%)
For i% = 0 To Max%
    WechselZeilen$(i%) = zeilen$(i%)
Next i%
WechselStart% = start%

frmWechsel.Show 1

start% = WechselStart%
WechselFenster% = WechselErg%

End Function


