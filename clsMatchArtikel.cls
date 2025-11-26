VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchArtikel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ANZMATCHKONVERTIERUNGEN = 24

Const LANGERNAME = 0
Const AEPundAVP = 1
Const PZNundAVP = 2
Const FESTBETRAG = 3

Const INI_SECTION = "Auswahl-Artikel"
Const INFO_SECTION = "Infobereich Auswahl-Artikel"

Dim MatchTaxeRec As Recordset

Dim TaxeAuswahlModus%

Dim MatchKlein$, MatchGross$, MatchGueltig$
Dim MatchPhonAlt$(ANZMATCHKONVERTIERUNGEN)
Dim MatchPhonNeu$(ANZMATCHKONVERTIERUNGEN)

Dim FabsErrf%
Dim FabsRecno&

Private Const DefErrModul = "clsMatchArtikel.cls"

Sub InitMatch()

ReDim MatchAnzeigeTyp$(2)
MatchAnzeigeTyp$(0) = "Taxe"
MatchAnzeigeTyp$(1) = "Taxe phonetisch"
MatchAnzeigeTyp$(2) = "Lagerartikel"

MatchKlein = "abcdefghijklmnopqrstuvwxyzäöüß,.-;:+*/ "
MatchGross = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜS........."
MatchGueltig = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ0123456789."
 
MatchPhonAlt$(0) = "J": MatchPhonNeu$(0) = "I"
MatchPhonAlt$(1) = "PH": MatchPhonNeu$(1) = "F"
MatchPhonAlt$(2) = "Y": MatchPhonNeu$(2) = "UE"
MatchPhonAlt$(3) = "Ä": MatchPhonNeu$(3) = "AE"
MatchPhonAlt$(4) = "Ö": MatchPhonNeu$(4) = "OE"
MatchPhonAlt$(5) = "Ü": MatchPhonNeu$(5) = "UE"
MatchPhonAlt$(6) = "SCH": MatchPhonNeu$(6) = "S"
MatchPhonAlt$(7) = "CH": MatchPhonNeu$(7) = "K"
MatchPhonAlt$(8) = "H": MatchPhonNeu$(8) = ""
MatchPhonAlt$(9) = "OE": MatchPhonNeu$(9) = "O"
MatchPhonAlt$(10) = "UE": MatchPhonNeu$(10) = "I"
MatchPhonAlt$(11) = "AI": MatchPhonNeu$(11) = "EI"
MatchPhonAlt$(12) = "AI": MatchPhonNeu$(12) = "EI"
MatchPhonAlt$(13) = "AE": MatchPhonNeu$(13) = "E"
MatchPhonAlt$(14) = "IE": MatchPhonNeu$(14) = "I"
MatchPhonAlt$(15) = "P": MatchPhonNeu$(15) = "B"
MatchPhonAlt$(16) = "T": MatchPhonNeu$(16) = "D"
MatchPhonAlt$(17) = "F": MatchPhonNeu$(17) = "V"
MatchPhonAlt$(18) = "V": MatchPhonNeu$(18) = "W"
MatchPhonAlt$(19) = "Z": MatchPhonNeu$(19) = "C"
MatchPhonAlt$(20) = "G": MatchPhonNeu$(20) = "K"
MatchPhonAlt$(21) = "C": MatchPhonNeu$(21) = "K"
MatchPhonAlt$(22) = "M": MatchPhonNeu$(22) = "N"
MatchPhonAlt$(23) = "O": MatchPhonNeu$(23) = "U"

'MatchModus% = TAXE_ALPHA
TaxeAuswahlModus% = AEPundAVP

End Sub

Function MakeMatchCode$(Name$)
Dim ch$, chAlt$, h$, h2$, zw$, i%, l%, p%, ind%

h$ = RTrim$(Name$)
l% = Len(h$)
For i% = 1 To l%
    ind% = InStr(MatchKlein$, Mid$(h$, i%, 1))
    If (ind% > 0) Then Mid$(h$, i%, 1) = Mid$(MatchGross$, ind%, 1)
Next i%

h2$ = ""
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    ind% = InStr(MatchGueltig$, ch$)
    If (ind% > 0) Then h2$ = h2$ + ch$
Next i%

h$ = h2$

For i% = 0 To (ANZMATCHKONVERTIERUNGEN - 1)
    p% = 2
    Do
        ind% = InStr(p%, h$, MatchPhonAlt(i%))
        If (ind% = 0) Then Exit Do
        If (Mid$(h$, ind% - 1, 1) <> ".") Then
            zw$ = Mid$(h$, ind% + Len(MatchPhonAlt(i%)))
            h$ = Left$(h$, ind% - 1) + MatchPhonNeu(i%) + zw$
        Else
            p% = p% + 1
        End If
    Loop
Next i%

l% = Len(h$)
h2$ = ""
chAlt$ = Chr$(1)
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    If (ch$ <> chAlt$) Then h2$ = h2$ + ch$
    chAlt$ = ch$
Next i%

h$ = h2$

MakeMatchCode$ = RTrim$(Left$(h2$ + Space$(10), 10))

End Function

Sub Umspeichern(buf$, pos%)
Dim ind%, ind2%, i%, z%, z1%, z2%
Dim zu!
Dim EK#, VK#
Dim h$, h2$, char$, zuza$, zusa$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

Select Case MatchModus%
    Case TAXE_ALPHA, TAXE_MATCH
        
        Ausgabe(pos%).pzn = MatchTaxeRec!pzn
        
        h$ = " "
        FabsErrf% = Zus1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
        If (FabsErrf% = 0) Then
            Ausgabe(pos%).ZusatzTextKz = 2
            h$ = "+"
        End If
        FabsErrf% = Besorgt1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
        If (FabsErrf% = 0) Then
            Ausgabe(pos%).BesorgtKz = 2
            If (Ausgabe(pos%).ZusatzTextKz = 2) Then
                h$ = "±"
            Else
                h$ = "-"
            End If
        End If
        Ausgabe(pos%).Name = h$ + vbTab + MachTaxeAuswahlName$(MatchTaxeRec)
        
        Ausgabe(pos%).LagerKz = 0
        Ausgabe(pos%).Bookmark = MatchTaxeRec.Bookmark
            
        FabsErrf% = Ass1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
        If (FabsErrf% = 0) Then Ausgabe(pos%).LagerKz = 2
        
            
            
    Case LAGER_MATCH
        Ausgabe(pos%).pzn = Ast1.pzn
        Ausgabe(pos%).Bookmark = Ast1.FabsRecno
        Ausgabe(pos%).key = clsfabs.LastKey(10)
        
        EK# = Ast1.AEP
'        Call DxToIEEEd(EK#)
        VK# = Ast1.AVP
'        Call DxToIEEEd(VK#)
        
        zuza$ = ""
        zu! = Ast1.zuza
        If (zu! < 0) Then zu! = zu! + 65536!
        '* 1 x shift right
        z% = (zu! / 2)
        For z1% = 1 To 5
          z2% = (z% Mod 8)
          '* 3 x shift right
          z% = z% \ 8
          zuza$ = Right$(Str$(z2%), 1) + zuza$
        Next z1%
        zuza$ = MachStammAuswahlZuzahlung$(zuza$)
        
        h$ = " "
        FabsErrf% = Zus1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
        If (FabsErrf% = 0) Then
            Ausgabe(pos%).ZusatzTextKz = 2
            h$ = "+"
        End If
        FabsErrf% = Besorgt1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
        If (FabsErrf% = 0) Then
            Ausgabe(pos%).BesorgtKz = 2
            If (Ausgabe(pos%).ZusatzTextKz = 2) Then
                h$ = "±"
            Else
                h$ = "-"
            End If
        End If
        
'        zusa$ = ast.vc: If (zusa$ <> "S") Then zusa$ = " "
'        zusa$ = zusa$ + ast.lic + ast.lac + ast.Rez + ast.kas
'        zusa$ = zusa$ + ast.kasz + ast.wg + Left$(ast.abl, 1)
'
'        h$ = h$ + vbTab + ast.kurz + vbTab + ast.meng + vbTab + ast.meh + vbTab
'        h$ = h$ + zuza$ + vbTab + ast.Herst + vbTab + Str$(EK#) + vbTab + Str$(VK#) + vbTab + zusa$
        
        zusa$ = Ast1.vc: If (zusa$ <> "S") Then zusa$ = " "
        zusa$ = zusa$ + Ast1.lic + Ast1.lac + Ast1.Rez + Ast1.kas
        zusa$ = zusa$ + Ast1.kasz + Ast1.wg + Left$(Ast1.abl, 1)
        
        h$ = h$ + vbTab + Ast1.kurz + vbTab + Ast1.meng + vbTab + Ast1.meh + vbTab
        h$ = h$ + zuza$ + vbTab + Ast1.herst + vbTab + Str$(EK#) + vbTab + Str$(VK#) + vbTab + zusa$
        
        Ausgabe(pos%).Name = h$
End Select
End Sub

Function MachStammAuswahlZuzahlung$(zuza$)
Dim ch$, ZuzWert!, i%, zz%, VglWert!, AVP!

ch$ = Left$(zuza$, 1)
'If (tRec!Warengruppe = 14) Then ch$ = "1"
If (ch$ = "9") Then
    MachStammAuswahlZuzahlung$ = "20%"
ElseIf (ch$ = "0") Then
    MachStammAuswahlZuzahlung$ = "kA"
ElseIf (ch$ = "1") Then
    MachStammAuswahlZuzahlung$ = "nb"
ElseIf (ch$ = "2") Then
    MachStammAuswahlZuzahlung$ = "ng"
Else
    ZuzWert! = 0
    For i% = 1 To 5
        zz% = Val(Mid$(zuza$, i%, 1))
        ZuzWert! = ZuzWert! + Para1.ZuzWert(zz%)
    Next i%
        
    MachStammAuswahlZuzahlung$ = Format$(ZuzWert!, "##")
End If

End Function


Function MachTaxeAuswahlName$(tRec As Recordset)
Dim i%, h$, h2$

If (tRec!SieheKz) Then
    h$ = tRec!Name
    For i% = 1 To frmMatchcode!flxarbeit(0).Cols - 1
        h$ = h$ + vbTab
    Next i%
Else
    h2$ = MachTaxeAuswahlZuzahlung$(tRec)
    h$ = tRec!Name + vbTab + tRec!menge + vbTab + tRec!Einheit + vbTab + h2$ + vbTab
    h2$ = ""
    If (tRec!ArtStatus = "S") Then h2$ = h2$ + "S"
    If ((tRec!AbgabeBest = 1) Or _
        (tRec!AbgabeBest = 5) Or _
        (tRec!AbgabeBest = 6) _
        ) Then h2$ = h2$ + "+"
    If (tRec!NegListe) Then h2$ = h2$ + "*"
    If (tRec!RufKz) Then h2$ = h2$ + "RW"
    If (tRec!OriginalPZN = tRec!pzn) Then
        h2$ = h2$ + "i"
    ElseIf (tRec!OriginalPZN > 0) Then
        h2$ = h2$ + "I"
    End If
    If (tRec!AlternativPackung > 0) Then h2$ = h2$ + "g"
        
    Select Case TaxeAuswahlModus%
        Case LANGERNAME
            h$ = h$ + h2$
        Case AEPundAVP
            h$ = h$ + tRec!HerstellerKB + vbTab + Format$(tRec!EK / 100, "0.00") + vbTab + Format$(tRec!VK / 100, "0.00") + vbTab + h2$
        Case PZNundAVP
            h$ = h$ + tRec!HerstellerKB + vbTab + Format$(tRec!pzn, "0000000") + vbTab + Format$(tRec!VK / 100, "0.00") + vbTab + h2$
        Case FESTBETRAG
            If (tRec!FESTBETRAG = 0&) Then
                h$ = h$ + vbTab + vbTab + Format$(tRec!VK / 100, "0.00") + vbTab + h2$
            Else
                h$ = h$ + Format$(tRec!FESTBETRAG / 100, "0.00") + vbTab + Format$((tRec!VK - tRec!FESTBETRAG) / 100, "0.00") + vbTab + Format$(tRec!VK / 100, "0.00") + vbTab + h2$
            End If
    End Select
End If

MachTaxeAuswahlName$ = h$

End Function

Function MachTaxeAuswahlZuzahlung$(tRec As Recordset)
Dim ch$, ZuzWert!, i%, zz%, VglWert!, AVP!

ch$ = Left$(tRec!Zuzahlung, 1)
If (tRec!Warengruppe = 14) Then ch$ = "1"
If (ch$ = "9") Then
    MachTaxeAuswahlZuzahlung$ = "20%"
ElseIf (ch$ = "0") Then
    MachTaxeAuswahlZuzahlung$ = "kA"
ElseIf (ch$ = "1") Then
    MachTaxeAuswahlZuzahlung$ = "nb"
ElseIf (ch$ = "2") Then
    MachTaxeAuswahlZuzahlung$ = "ng"
Else
    ZuzWert! = 0
    For i% = 1 To 5
        zz% = Val(Mid$(tRec!Zuzahlung, i%, 1))
        ZuzWert! = ZuzWert! + Para1.ZuzWert(zz%)
    Next i%
        
    AVP! = tRec!VK / 100
    If ((tRec!FestKz) And (tRec!FESTBETRAG < AVP!)) Then
        VglWert! = tRec!FESTBETRAG
    Else
        VglWert! = AVP!
    End If
        
    If ((VglWert! > 0!) And (ZuzWert! > VglWert!)) Then
        MachTaxeAuswahlZuzahlung$ = "VK"
    Else
        MachTaxeAuswahlZuzahlung$ = Format$(ZuzWert!, "##")
    End If
End If

End Function

Function SuchWeiter%(ind%, Typ%)
Dim AltRecno&
Dim h$

SuchWeiter% = False

Select Case MatchModus%
    Case TAXE_ALPHA, TAXE_MATCH
        MatchTaxeRec.Bookmark = Ausgabe(ind%).Bookmark
        If (Typ%) Then
            MatchTaxeRec.MoveNext
            If (MatchTaxeRec.EOF = False) Then SuchWeiter% = True
        Else
            MatchTaxeRec.MovePrevious
            If (MatchTaxeRec.BOF = False) Then SuchWeiter% = True
        End If
    Case LAGER_MATCH
        AltRecno& = Ausgabe(ind%).Bookmark
        If (Typ%) Then
            FabsErrf% = Ast1.IndexNext(1, AltRecno&, RTrim$(Ausgabe(ind%).key), FabsRecno&)
        Else
            FabsErrf% = Ast1.IndexPrevious(1, AltRecno&, RTrim$(Ausgabe(ind%).key), FabsRecno&)
        End If
        If (FabsErrf% = 0) Then
            Ast1.GetRecord (FabsRecno& + 1)
            SuchWeiter% = True
        End If
End Select

End Function

Function SuchArtikel%(Such$, iArbeitAnzzeilen%)
Dim i%, l%, ind%, TextSuche%, IstPzn%, LeftLen%, InhMengeFlag%
Dim InhMenge&
Dim SQLStr$, nSuch$, ch$
Dim iBookmark As Variant

SuchArtikel% = True

OrgSuch$ = Such$

If (MatchModus% = LAGER_MATCH) Then
    Such$ = MakeMatchCode$(Such$)
'    FabsErrf% = clsfabs.Cmnd("G\1\18\" + Such$, FabsRecno&)
    FabsErrf% = Ast1.IndexGeneric(1, Such$, FabsRecno&)
    If ((FabsErrf% = 0) Or (FabsErrf% = 12) Or (FabsErrf% = 13) Or (FabsErrf% = 15)) Then
        Ast1.GetRecord (FabsRecno& + 1)
        
        Call Umspeichern("", 0)
        Call MachAuswahlGrid(iArbeitAnzzeilen%)
        Call frmMatchcode.MachMaske
        With frmMatchcode.flxarbeit(0)
            .row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    Else
        SuchArtikel% = False
    End If
    Exit Function
End If


If (Such$ = "") Then Such$ = Space$(10)
l% = Len(Such$)
If ((l% >= 10) And (Left$(Such$, 1) = "*") And (Mid$(Such$, l%, 1) = "*")) Then
    nSuch$ = ""
    For i% = 1 To l%
        ch$ = Mid$(Such$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            nSuch$ = nSuch$ + ch$
        End If
    Next i%
    Such$ = nSuch$
End If

IstPzn% = True
TextSuche% = False
For i% = 1 To l%
    ch$ = Mid$(Such$, i%, 1)
    If (InStr("0123456789", ch$) = 0) Then
        IstPzn% = False
'        Exit For
    End If
    If (InStr(",:*?<>", ch$) > 0) Then
        TextSuche% = True
    End If
Next i%

If ((TextSuche% = False) And (IstPzn% = False)) Then
    If (MatchModus% = TAXE_ALPHA) Then
        LeftLen% = l%
        InhMenge& = 0
        For i% = 1 To l%
            ch$ = Mid$(Such$, i%, 1)
            If ((ch$ >= "0") And (ch$ <= "9")) Then
                If (InhMengeFlag% = False) Then
                    LeftLen% = i% - 1
                    InhMengeFlag% = True
                End If
                InhMenge& = (InhMenge& * 10) + (Asc(ch$) - 48)
            ElseIf (InhMengeFlag% = True) Then
                Exit For
            End If
        Next i%
        Such$ = Left$(Such$, LeftLen%)
    ElseIf (MatchModus% = TAXE_MATCH) Then
        Such$ = MakeMatchCode$(Such$)
    End If
End If

If ((TextSuche% = True) Or (IstPzn% = True)) Then
'    Call StartAnimation(frmAction.aniSuchen)
    If (IstPzn% = True) Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Such$
    ElseIf (MatchModus% = TAXE_ALPHA) Then
        If (Left$(Such$, 1) = "<") Then Such$ = Mid$(Such$, 2) + "*"
        ind% = InStr(Such$, ",")
        If (ind% > 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE NAME LIKE '*" + Left$(Such$, ind% - 1) + _
            "*' AND NAME LIKE '*" + Mid$(Such$, ind% + 1) + "*'"
        Else
            ind% = InStr(Such$, ":")
            If (ind% > 0) Then
                SQLStr$ = "SELECT * FROM TAXE WHERE NAME LIKE '*" + Left$(Such$, ind% - 1) + _
                "*' OR NAME LIKE '*" + Mid$(Such$, ind% + 1) + "*'"
            Else
                ch$ = Right$(Such$, 1)
                If (ch$ <> "*") Then
                    If (ch$ = ">") Then
                        Such$ = "*" + Left$(Such$, Len(Such$) - 1)
                    Else
                        Such$ = Such$ + "*"
                    End If
                End If
                SQLStr$ = "SELECT * FROM TAXE WHERE NAME LIKE '" + Such$ + "'"
            End If
        End If
    Else
        SQLStr$ = "SELECT * FROM TAXE WHERE MATCHCODE LIKE """ + Such$ + "*"""
    End If

'    frmAction!picBack(0).MousePointer = vbHourglass
'    DoEvents
    Set MatchTaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    frmAction!picBack(0).MousePointer = vbDefault

    If (MatchTaxeRec.EOF) Then
        SuchArtikel% = False
    Else
'        Call Umspeichern("", 0)
'        TaxeRec.MoveLast
'        ListeSize% = TaxeRec.RecordCount
    End If
'    Call StopAnimation(frmAction.aniSuchen)
Else
    Set MatchTaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
    If (MatchModus% = TAXE_ALPHA) Then
        MatchTaxeRec.index = "KurzName"
        MatchTaxeRec.Seek ">=", Such$, InhMenge&
        If (MatchTaxeRec.NoMatch = False) Then
            iBookmark = MatchTaxeRec.Bookmark
            MatchTaxeRec.index = "Name"
            MatchTaxeRec.Bookmark = iBookmark
        End If
    Else
        MatchTaxeRec.index = "MatchCode"
        MatchTaxeRec.Seek ">=", Such$
    End If
    If (MatchTaxeRec.NoMatch = True) Then
        SuchArtikel% = False
    Else
'        Call Umspeichern("", 0)
    End If
End If

If (SuchArtikel% = True) Then
    Call Umspeichern("", 0)
    Call MachAuswahlGrid(iArbeitAnzzeilen%)
    Call frmMatchcode.MachMaske
    With frmMatchcode.flxarbeit(0)
        .row = 1
        .ColSel = .Cols - 1
        .SetFocus
    End With
End If

End Function

Sub MachAuswahlGrid(iArbeitAnzzeilen%)
Dim spBreite%, i%
Dim aFont$

With frmMatchcode.flxarbeit(0)
    .Rows = iArbeitAnzzeilen%
    .FixedRows = 1
            
    Select Case TaxeAuswahlModus%
        Case LANGERNAME
            .Cols = 6
            .FormatString = "|<Name|>Menge|<Eh|Zu|<Kz"
            frmMatchcode.Font.Bold = True
            .ColWidth(0) = frmMatchcode.TextWidth("X")
            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
            .ColWidth(4) = frmMatchcode.TextWidth("XX")
            .ColWidth(5) = frmMatchcode.TextWidth("XXX")
            frmMatchcode.Font.Bold = False
        Case AEPundAVP, PZNundAVP
            If (MatchModus% = LAGER_MATCH) Then
                .Cols = 9   ' 16
                .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
                frmMatchcode.Font.Bold = True
                .ColWidth(0) = frmMatchcode.TextWidth("X")
                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
                .ColWidth(4) = frmMatchcode.TextWidth("XX")
                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
                aFont$ = frmMatchcode.Font.Name
                frmMatchcode.Font.Name = "Courier New"
                frmMatchcode.Font.Bold = True
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
                frmMatchcode.Font.Name = aFont$
                frmMatchcode.Font.Bold = False
            Else
                .Cols = 9
                If (TaxeAuswahlModus% = AEPundAVP) Then
                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
                Else
                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|<PZN|>VK|<Kz"
                End If
                frmMatchcode.Font.Bold = True
                .ColWidth(0) = frmMatchcode.TextWidth("X")
                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
                .ColWidth(4) = frmMatchcode.TextWidth("XX")
                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXX")
                frmMatchcode.Font.Bold = False
            End If
        Case FESTBETRAG
            .Cols = 9
            .FormatString = "|<Name|>Menge|<Eh|^Zu|>FB|>DIFF|>VK|<Kz"
            frmMatchcode.Font.Bold = True
            .ColWidth(0) = frmMatchcode.TextWidth("X")
            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
            .ColWidth(4) = frmMatchcode.TextWidth("XX")
            .ColWidth(5) = frmMatchcode.TextWidth("99999.99")
            .ColWidth(6) = frmMatchcode.TextWidth("999.99")
            .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
            .ColWidth(8) = frmMatchcode.TextWidth("XXX")
            frmMatchcode.Font.Bold = False
    End Select
            
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (i <> 1) Then
            .ColWidth(i%) = .ColWidth(i%) + frmMatchcode.TextWidth("X")
            spBreite% = spBreite% + .ColWidth(i%)
        End If
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(1) = .Width - spBreite%
End With
End Sub


Public Property Get IniSection$()
IniSection$ = INI_SECTION
End Property

Public Property Get InfoSection$()
InfoSection$ = INFO_SECTION
End Property

