Attribute VB_Name = "modGhAngebote"
Option Explicit

Type LocalAngeboteStruct
    gh As Byte
    st As String * 1
    bm As Byte
    mp As Long
    AepKalk As Double
End Type

Public AngebotPzn$
Public AngebotModus%
Public AngebotInd%
Public AngebotAuswahl%
Public AngebotAep#
Public AngebotLief%
Public AngebotBm%
Public AngebotNm%

Dim zr!
Dim bm%, nr%

Public LocalAngebotRec(50) As LocalAngeboteStruct
Public GhAngebotMax&

Dim AnzLocalAngebote%


Dim NrWert#, BarRab#, LaKo#, PeKo#, Manko#, StaffelAb#, AEP#

Dim BMopt!, BMast%

Dim FabsRecno&
Dim FabsErrf%

Private Const DefErrModul = "wangebote.bas"

Sub Angebote()
Dim TaxeAep#, AnfSatz&, ssatz&, asatz&, SQLStr$
Dim pzn$

pzn$ = AngebotPzn$

GhAngebotMax& = clsAngebote1.DateiLen / 34 - 1
AnfSatz& = AngebotSuchen&(pzn$)

If (AngebotAuswahl% = False) Then
    If (AnfSatz& > 0) Then
        AngebotInd% = 1
    Else
        AngebotInd% = 0
    End If
End If
    
If (AnfSatz& > 0) Then
    ssatz& = 0
    Ass1.pzn = String$(7, 0)
    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ssatz& = FabsRecno&
        Ass1.GetRecord (FabsRecno& + 1)
    End If
    
    asatz& = 0
    Ast1.pzn = String$(7, 0)
    FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        asatz& = FabsRecno&
        Ast1.GetRecord (FabsRecno& + 1)
    End If
    
  
    If AngebotAuswahl% Or (AngebotAuswahl% = 0 And AngebotLief% = 0) Then
       TaxeAep# = 0#

        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            TaxeAep# = TaxeRec!EK / 100
            If (asatz& <= 0) Then
                Call clsOpTool.Taxe2ast(pzn$)
            End If
        End If

        Call AngebotSelect(AnfSatz&, TaxeAep#)
    End If
End If

End Sub

Function AngebotSuchen&(SuchPzn$)
Dim Angebote%
Dim lAnzAngebote&, von&, bis&, index&, Such$, ret&
Dim s$

If (Dir$("\user\gh-anbot.dat") = "") Then
    AngebotSuchen& = -1&
    Exit Function
End If

ret& = -1&
If (SuchPzn$ = "9999999") Then Exit Function

'Angebote% = FileOpen%("\user\gh-anbot.dat", "R")
'lAnzAngebote& = LOF(GHANBOT%) / 34 - 1

von& = 1&
bis& = GhAngebotMax&  'lAnzAngebote&
'Such$ = Format(SuchPzn&, "0000000")

Do While (von& <= bis&)
    index& = (von& + bis&) \ 2
    Call clsAngebote1.GetRecord(index& + 1)
    s$ = clsAngebote1.pzn
    If (SuchPzn$ = s$) Then
        ret& = index&
        Exit Do
    ElseIf (SuchPzn$ < s$) Then
        bis& = index& - 1
    Else
        von& = index& + 1
    End If
Loop

If (ret& >= 1) Then
    Do
        ret& = ret& - 1
        Call clsAngebote1.GetRecord(ret& + 1)
        s$ = clsAngebote1.pzn
        If (SuchPzn$ <> s$) Then
            Exit Do
        End If
    Loop
    ret& = ret& + 1
End If

'Close #Angebote%

AngebotSuchen& = ret&

End Function

Sub AngebotSelect(AnfSatz&, taxAEP#)
Dim i%, j%, k%, gh%, BMorg%, GesMenge%
Dim AEPorg#, AepKalk#, gspart#, prozgspart#, minaep#, aep1#, aep2#
Dim satz&, mp&, MPorg&
Dim pzn$, st$, h$, lKurz$
Dim iLocalAngebotRec As LocalAngeboteStruct

If (Ass1.pzn <> String$(7, 0)) Then
    BMopt! = Ass1.opt
    BMast% = Ass1.bm
Else
    BMopt! = 0
    BMast% = 0
End If
If (Ast1.pzn <> String$(7, 0)) Then
    AEPorg# = Ast1.AEP
Else
    AEPorg# = 0#
End If
'Call DxToIEEEd(AEPorg#)

If BMopt! > 3200 Then BMopt! = 3200
If BMopt! <= 0 Then BMopt! = 1      '2.55

satz& = AnfSatz&
AnzLocalAngebote% = 0

Call clsAngebote1.GetRecord(satz& + 1)
pzn$ = clsAngebote1.pzn
While (pzn$ = clsAngebote1.pzn) And (satz& <= GhAngebotMax&)
    gh% = clsAngebote1.gh
    st$ = clsAngebote1.st
    For i% = 0 To 4
        bm% = clsAngebote1.bm(i%): BMorg% = bm% '2.84
        mp& = clsAngebote1.mp(i%): MPorg& = mp&

        '* Naturalrabatt
        If st$ = "M" And bm% > 0 And mp& > 0 Then
            'kleinstes Angebot unterhalb der optimalen Bestellmenge
            While bm% < BMopt!
              bm% = bm% + BMorg%
              mp& = mp& + MPorg&
            Wend
            If bm% <> BMorg% And bm% > BMopt! Then
              'hoppla; eins zu weit, da erst bei größer gleich abgebrochen wird
              bm% = bm% - BMorg%
              mp& = mp& - MPorg&
            End If
        End If

        '* Preisrabatt
        If st$ = "P" And bm% > 0 And mp& > 0 Then
            If bm% < BMopt! Then bm% = Int(BMopt!)
        End If

        While bm% > 0 And mp& > 0
            If AnzLocalAngebote% < 50 Then
                LocalAngebotRec(AnzLocalAngebote%).gh = gh%
                LocalAngebotRec(AnzLocalAngebote%).st = st$
                LocalAngebotRec(AnzLocalAngebote%).bm = bm%
                LocalAngebotRec(AnzLocalAngebote%).mp = mp&
                AnzLocalAngebote% = AnzLocalAngebote% + 1
            End If
            If st$ = "M" And bm% < BMopt! Then
                'wenn das Mengenangebot kleiner ist als die optimale
                'Bestellmenge versuchen wir noch das Angebot darüber
                bm% = bm% + BMorg%
                mp& = mp& + MPorg&
            Else
                bm% = 0: mp& = 0
            End If
        Wend
    Next i%

    satz& = satz& + 1
    Call clsAngebote1.GetRecord(satz& + 1)
Wend

minaep# = 9999999999#
For i% = 0 To (AnzLocalAngebote% - 1)
    Call AngebotAuswerten(i%, AEPorg#, AepKalk#, gspart#, prozgspart#)
    LocalAngebotRec(i%).AepKalk = AepKalk#
    If AepKalk# > 0 And AepKalk# < minaep# And AepKalk# < AEPorg# And BMast% > 0 Then
        minaep# = AepKalk#
        If (AngebotAuswahl% = False) Then
            AngebotInd% = (i% + 1) * (-1)
            AngebotLief% = LocalAngebotRec(i%).gh
            AngebotBm% = LocalAngebotRec(i%).bm
            AngebotNm% = 0
            If LocalAngebotRec(i%).st = "M" Then AngebotNm% = LocalAngebotRec(i%).mp
            AngebotAep# = AepKalk#
        End If
    End If
Next i%

If (AngebotAuswahl%) Then
    'sortieren nach AEPkalk
    For i% = 0 To AnzLocalAngebote% - 2
      aep1# = LocalAngebotRec(i%).AepKalk
      For j% = i% + 1 To AnzLocalAngebote% - 1
        aep2# = LocalAngebotRec(j%).AepKalk
        If aep1# <= 0# Or aep1# > aep2# Then
          aep1# = aep2#
          iLocalAngebotRec = LocalAngebotRec(i%)
          LocalAngebotRec(i%) = LocalAngebotRec(j%)
          LocalAngebotRec(j%) = iLocalAngebotRec
        End If
      Next j%
    Next i%

  For i% = 0 To AnzLocalAngebote% - 1
        Call AngebotAuswerten(i%, AEPorg#, AepKalk#, gspart#, prozgspart#)

        lKurz$ = ""
        gh% = LocalAngebotRec(i%).gh
        If (gh% > 0) And (gh% < 200) Then
            Lif1.GetRecord (gh% + 1)
            h$ = Trim$(Lif1.kurz)
            If (h$ = String$(Len(h$), 0)) Then h$ = ""
            If (h$ = "") Then
                h$ = "(" + Str$(gh%) + ")"
            Else
                Call OemToChar(h$, h$)
            End If
            lKurz$ = h$
        End If

        With frmAngebote.flxAngebote
            GesMenge% = bm% + nr%
'            .Cols = .Cols + 1
            .TextMatrix(1, i% + 1) = lKurz$
            .TextMatrix(2, i% + 1) = Str$(bm%) + " Stk"
            .TextMatrix(3, i% + 1) = Format(AEPorg#, "0.00")
            .TextMatrix(4, i% + 1) = Format(-AEPorg# * (zr!) / 100#, "0.00")
            .TextMatrix(5, i% + 1) = Format(AEP#, "0.00")
            .TextMatrix(6, i% + 1) = Str$(nr%) + Format(NrWert# / GesMenge%, "   (-0.00)")
            .TextMatrix(7, i% + 1) = Format(LaKo# / GesMenge%, "0.00")
            .TextMatrix(8, i% + 1) = Format(PeKo# / GesMenge%, "0.00")
            .TextMatrix(9, i% + 1) = Format(StaffelAb# / GesMenge%, "0.00")
            .TextMatrix(10, i% + 1) = Format(AepKalk#, "0.00")
            .TextMatrix(11, i% + 1) = Format(gspart#, "0.00")
            .TextMatrix(12, i% + 1) = Format(prozgspart#, "0")

            If (i% = AngebotInd%) Then
                .col = i% + 1
                For k% = 1 To .Rows - 1
                    .row = k%
                    .CellFontBold = True
                Next k%
            End If
        End With
    Next i%

    With frmAngebote.flxAngebote
        .Cols = AnzLocalAngebote% + 1
        If (AngebotInd% >= 0) Then
            .col = AngebotInd% + 1
        Else
            .col = 1
        End If
        .row = 1
        .RowSel = .Rows - 1
    End With

    With frmAngebote
        .lblAngeboteWert(0) = BMopt!
        .lblAngeboteWert(1) = AEPorg#
        .lblAngeboteWert(2) = taxAEP#
    End With

End If

End Sub

Static Sub AngebotAuswerten(angebot%, AEPorg#, AepKalk#, gspart#, prozgspart#)
Dim GesMenge%

nr% = 0: zr! = 0
bm% = LocalAngebotRec(angebot%).bm
If LocalAngebotRec(angebot%).st = "P" Then
  If AEPorg# > 0.009 Then zr! = 100 - LocalAngebotRec(angebot%).mp / AEPorg#
  AEP# = LocalAngebotRec(angebot%).mp / 100#
Else
  nr% = LocalAngebotRec(angebot%).mp
  AEP# = AEPorg#
End If
Call AngebotRechnen(bm%, nr%, zr!, AEP#)
GesMenge% = bm% + nr%
AepKalk# = AEP# + Manko# + (StaffelAb# + LaKo# + PeKo# - NrWert#) / GesMenge%
prozgspart# = 0#
If AEPorg# > 0.009 Then
  gspart# = (AEPorg# - AepKalk#) * GesMenge%
  prozgspart# = (AEPorg# - AepKalk#) / AEPorg# * 100#
End If

End Sub

Static Sub AngebotRechnen(bm%, nr%, zr!, AEP#)
Dim wert#, lx#, fr#, LaZi#
Dim zuwenig!, zuviel!, WieLange!
Dim GesMenge%

fr# = 0#        'FakturenRabatt
LaKo# = 0#
PeKo# = 0#
LaZi# = 0#

If bm% <> 0 Or nr% <> 0 Then

  If BMopt! <= 0 Then BMopt! = BMast%

  GesMenge% = bm% + nr%

  '* Bestellwert errechnen
  wert# = CDbl(bm%) * AEP#

  '* Naturalrabattwert (mit Faktor)
  NrWert# = CDbl(nr%) * AEP# * Para1.FakNatu

  '* Zeilen-Barrabattwert (mit Faktor)
  BarRab# = ((CDbl(zr!) + fr#) / 100#) * CDbl(bm%) * AEP# * Para1.FakBar

  'ist optimierung vorhanden ?  (-1)=nein
  If BMopt! > 0 Then

    '* zuviel bestellt
    If GesMenge% > BMopt! And BMopt! > 0 Then
      zuviel! = CSng(GesMenge%) - BMopt!
      WieLange! = (zuviel! / BMopt!) * Para1.BestellPeriode
      LaKo# = WieLange! / 360# * (Para1.Lagerkosten) / 100# * AEP# * zuviel!
      PeKo# = WieLange! / 360# * (Para1.PersonalKosten) / 100# * AEP# * zuviel!
      LaZi# = LaKo# + PeKo#
    End If

    '* zuwenig bestellt
    If GesMenge% < BMopt! And BMopt! > 0 Then
      zuwenig! = BMopt! - CSng(GesMenge%)
      WieLange! = (zuwenig! / BMopt!) * Para1.BestellPeriode
      lx# = WieLange! / 360# * (Para1.PersonalKosten) / 100# * AEP# * zuwenig!
      LaZi# = LaZi# + lx#
      PeKo# = PeKo# + lx#
    End If
  End If
  Manko# = (100# - Para1.Tfaktor) * AEP# / 100#
  StaffelAb# = 0#
  If BMast% > 0 Then
    If (bm% + nr%) > BMast% Then
      StaffelAb# = ((bm% + nr%) - BMast%) / BMast% * Para1.PersonalKosten / 100# * AEP#
    End If
  End If
End If

End Sub

