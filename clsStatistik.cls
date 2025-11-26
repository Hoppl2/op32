VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 256

Private Type StatistikErstSatz
    erstmax As Integer
    erstrest As String * 254
End Type
Private Type Stat1
    datum As String * 2
    menge As Integer
End Type
Private Type llief1
    lief As Integer
    datum As String * 2
    lm As Integer
    nr As Integer
End Type
Private Type Verk1
    datum As String * 2 'Integer
    menge As Integer
End Type
Private Type Teinzel1
    teplatz As Byte
    telager As Byte
    temm As Byte
End Type
Private Type StatistikStruct
    pzn As String * 7
    lief As Byte
    halt As Byte
    lld As String * 2
    llm As Integer
    abl1 As String * 2
    abl2 As String * 2
    vbm As Integer
    bm As Integer
    mm As Integer
    lag As Integer
    opt As Integer
    stat(17) As Stat1
    pp As String * 1
    poslag As Integer
    tplatz As Integer
    tlager As Integer
    tmm As Integer
    EAN As String * 13
    quart(3) As Integer
    llief(4) As llief1
    verk(9) As Verk1
    quarttag As String * 2
    lagbew As Byte
    fplatz As Integer
    flager As Integer
    fmm As Integer
    vmm As Integer
    teinzel(3) As Teinzel1
    anlage As String * 2
    pk As String * 2
    rest2 As String * 20
End Type
Dim assErst As StatistikErstSatz
Dim ass As StatistikStruct


Dim iName$(2)
Dim iHandle%(2)
Dim iAnzIndex%

Dim dKurz$
Dim dBezeichnung$
Dim dName$
Dim dZugriff$
Dim dModus$
Dim dSatzLen%
Dim dHandle%
Dim dAnzFields%

Dim fKurz$()
Dim fBezeichnung$()
Dim fStart%()
Dim fLaenge%()
Dim fTyp$()
Dim fKonv%()

Dim FabsErrf%
Dim iFabsRecno&

Dim AktData As String * SATZLEN


Private Const DefErrModul = "clsStatistik.cls"

Private Sub Class_Initialize()

dKurz$ = "ASS"
dBezeichnung$ = "Statistik"
dName$ = "astatis.tax"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(ass)

dAnzFields% = 120
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pharmazentralnummer"
fStart%(0) = 1
fLaenge%(0) = 7
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "POSLAG"
fBezeichnung$(1) = "Lagerstand POS"
fStart%(1) = 101
fLaenge%(1) = 2
fTyp$(1) = "N"
fKonv%(1) = 5

fKurz$(2) = "OPT"
fBezeichnung$(2) = "Optimale Bestellmenge"
fStart%(2) = 26
fLaenge%(2) = 2
fTyp$(2) = "N"
fKonv%(2) = 6

fKurz$(3) = "LIEF"
fBezeichnung$(3) = "üblicher Lieferant"
fStart%(3) = 8
fLaenge%(3) = 1
fTyp$(3) = "N"
fKonv%(3) = 1

fKurz$(4) = "BM"
fBezeichnung$(4) = "Bestellmenge"
fStart%(4) = 20
fLaenge%(4) = 2
fTyp$(4) = "N"
fKonv%(4) = 5

fKurz$(5) = "MM"
fBezeichnung$(5) = "Mindestmenge"
fStart%(5) = 22
fLaenge%(5) = 2
fTyp$(5) = "N"
fKonv%(5) = 5

fKurz$(6) = "VBM"
fBezeichnung$(6) = "Vorgeschriebene Bestellmenge"
fStart%(6) = 18
fLaenge%(6) = 2
fTyp$(6) = "N"
fKonv%(6) = 5

fKurz$(7) = "VMM"
fBezeichnung$(7) = "Vorgeschriebene Mindestmenge"
fStart%(7) = 219
fLaenge%(7) = 2
fTyp$(7) = "N"
fKonv%(7) = 5

fKurz$(8) = "HALT"
fBezeichnung$(8) = "2 abgebbar/Interne Streichung"
fStart%(8) = 9
fLaenge%(8) = 1
fTyp$(8) = "N"
fKonv%(8) = 1

fKurz$(9) = "LLD"
fBezeichnung$(9) = "Datum der letzten Lieferung"
fStart%(9) = 10
fLaenge%(9) = 2
fTyp$(9) = "D"
fKonv%(9) = 10

fKurz$(10) = "LLM"
fBezeichnung$(10) = "Letzte Liefermenge"
fStart%(10) = 12
fLaenge%(10) = 2
fTyp$(10) = "N"
fKonv%(10) = 5

fKurz$(11) = "ABL1"
fBezeichnung$(11) = "Ablaufdatum-1"
fStart%(11) = 14
fLaenge%(11) = 2
fTyp$(11) = "D"
fKonv%(11) = 10

fKurz$(12) = "ABL2"
fBezeichnung$(12) = "Ablaufdatum-2"
fStart%(12) = 16
fLaenge%(12) = 2
fTyp$(12) = "D"
fKonv%(12) = 10

fKurz$(13) = "LAG"
fBezeichnung$(13) = "Liefermenge alt"
fStart%(13) = 24
fLaenge%(13) = 2
fTyp$(13) = "N"
fKonv%(13) = 5

fKurz$(14) = "PP"
fBezeichnung$(14) = "Etikettenart"
fStart%(14) = 100
fLaenge%(14) = 1
fTyp$(14) = "T"
fKonv%(14) = 0

fKurz$(15) = "EAN"
fBezeichnung$(15) = "EAN-Code"
fStart%(15) = 109
fLaenge%(15) = 13
fTyp$(15) = "T"
fKonv%(15) = 0

fKurz$(16) = "REST"
fBezeichnung$(16) = "Rest"
fStart%(16) = 237
fLaenge%(16) = 20
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "ANLAGE"
fBezeichnung$(17) = "Datum der Anlage"
fStart%(17) = 233
fLaenge%(17) = 2
fTyp$(17) = "D"
fKonv%(17) = 10


fKurz$(18) = "LLief(4)"
fBezeichnung$(18) = "Lieferung 1 (jüngste): Lieferant"
fStart%(18) = 162
fLaenge%(18) = 2
fTyp$(18) = "N"
fKonv%(18) = 5

fKurz$(19) = "LLDat(4)"
fBezeichnung$(19) = "Lieferung 1: Datum"
fStart%(19) = 164
fLaenge%(19) = 2
fTyp$(19) = "D"
fKonv%(19) = 10

fKurz$(20) = "LLRM(4)"
fBezeichnung$(20) = "Lieferung 1: RM"
fStart%(20) = 166
fLaenge%(20) = 2
fTyp$(20) = "N"
fKonv%(20) = 5

fKurz$(21) = "LLNM(4)"
fBezeichnung$(21) = "Lieferung 1: NM"
fStart%(21) = 168
fLaenge%(21) = 2
fTyp$(21) = "N"
fKonv%(21) = 5

fKurz$(22) = "LLief(3)"
fBezeichnung$(22) = "Lieferung 2: Lieferant"
fStart%(22) = 154
fLaenge%(22) = 2
fTyp$(22) = "N"
fKonv%(22) = 5

fKurz$(23) = "LLDat(3)"
fBezeichnung$(23) = "Lieferung 2: Datum"
fStart%(23) = 156
fLaenge%(23) = 2
fTyp$(23) = "D"
fKonv%(23) = 10

fKurz$(24) = "LLRM(3)"
fBezeichnung$(24) = "Lieferung 2: RM"
fStart%(24) = 158
fLaenge%(24) = 2
fTyp$(24) = "N"
fKonv%(24) = 5

fKurz$(25) = "LLNM(3)"
fBezeichnung$(25) = "Lieferung 2: NM"
fStart%(25) = 160
fLaenge%(25) = 2
fTyp$(25) = "N"
fKonv%(25) = 5

fKurz$(26) = "LLief(2)"
fBezeichnung$(26) = "Lieferung 3: Lieferant"
fStart%(26) = 146
fLaenge%(26) = 2
fTyp$(26) = "N"
fKonv%(26) = 5

fKurz$(27) = "LLDat(2)"
fBezeichnung$(27) = "Lieferung 3: Datum"
fStart%(27) = 148
fLaenge%(27) = 2
fTyp$(27) = "D"
fKonv%(27) = 10

fKurz$(28) = "LLRM(2)"
fBezeichnung$(28) = "Lieferung 3: RM"
fStart%(28) = 150
fLaenge%(28) = 2
fTyp$(28) = "N"
fKonv%(28) = 5

fKurz$(29) = "LLNM(2)"
fBezeichnung$(29) = "Lieferung 3: NM"
fStart%(29) = 152
fLaenge%(29) = 2
fTyp$(29) = "N"
fKonv%(29) = 5

fKurz$(30) = "LLief(1)"
fBezeichnung$(30) = "Lieferung 4: Lieferant"
fStart%(30) = 138
fLaenge%(30) = 2
fTyp$(30) = "N"
fKonv%(30) = 5

fKurz$(31) = "LLDat(1)"
fBezeichnung$(31) = "Lieferung 4: Datum"
fStart%(31) = 140
fLaenge%(31) = 2
fTyp$(31) = "D"
fKonv%(31) = 10

fKurz$(32) = "LLRM(1)"
fBezeichnung$(32) = "Lieferung 4: RM"
fStart%(32) = 142
fLaenge%(32) = 2
fTyp$(32) = "N"
fKonv%(32) = 5

fKurz$(33) = "LLNM(1)"
fBezeichnung$(33) = "Lieferung 4: NM"
fStart%(33) = 144
fLaenge%(33) = 2
fTyp$(33) = "N"
fKonv%(33) = 5

fKurz$(34) = "LLief(0)"
fBezeichnung$(34) = "Lieferung 5: Lieferant"
fStart%(34) = 130
fLaenge%(34) = 2
fTyp$(34) = "N"
fKonv%(34) = 5

fKurz$(35) = "LLDat(0)"
fBezeichnung$(35) = "Lieferung 5: Datum"
fStart%(35) = 132
fLaenge%(35) = 2
fTyp$(35) = "D"
fKonv%(35) = 10

fKurz$(36) = "LLRM(0)"
fBezeichnung$(36) = "Lieferung 5: RM"
fStart%(36) = 134
fLaenge%(36) = 2
fTyp$(36) = "N"
fKonv%(36) = 5

fKurz$(37) = "LLNM(0)"
fBezeichnung$(37) = "Lieferung 5: NM"
fStart%(37) = 136
fLaenge%(37) = 2
fTyp$(37) = "N"
fKonv%(37) = 5


fKurz$(38) = "TPLATZ"
fBezeichnung$(38) = "HV Gesamt: Max. Lagerstand"
fStart%(38) = 103
fLaenge%(38) = 2
fTyp$(38) = "N"
fKonv%(38) = 5

fKurz$(39) = "TLAGER"
fBezeichnung$(39) = "HV Gesamt: Akt. Lagerstand"
fStart%(39) = 105
fLaenge%(39) = 2
fTyp$(39) = "N"
fKonv%(39) = 5

fKurz$(40) = "TMM"
fBezeichnung$(40) = "HV Gesamt: Min. Lagerstand"
fStart%(40) = 107
fLaenge%(40) = 2
fTyp$(40) = "N"
fKonv%(40) = 5

fKurz$(41) = "FPLATZ"
fBezeichnung$(41) = "Filiale:   Max. Lagerstand"
fStart%(41) = 213
fLaenge%(41) = 2
fTyp$(41) = "N"
fKonv%(41) = 5

fKurz$(42) = "FLAGER"
fBezeichnung$(42) = "Filiale:   Akt. Lagerstand"
fStart%(42) = 215
fLaenge%(42) = 2
fTyp$(42) = "N"
fKonv%(42) = 5

fKurz$(43) = "FMM"
fBezeichnung$(43) = "Filiale:   Min. Lagerstand"
fStart%(43) = 217
fLaenge%(43) = 2
fTyp$(43) = "N"
fKonv%(43) = 5


fKurz$(44) = "Max(0)"
fBezeichnung$(44) = "Lager A:   Max. Lagerstand"
fStart%(44) = 221
fLaenge%(44) = 1
fTyp$(44) = "N"
fKonv%(44) = 1

fKurz$(45) = "Akt(0)"
fBezeichnung$(45) = "Lager A:   Akt. Lagerstand"
fStart%(45) = 222
fLaenge%(45) = 1
fTyp$(45) = "N"
fKonv%(45) = 1

fKurz$(46) = "Min(0)"
fBezeichnung$(46) = "Lager A:   Min. Lagerstand"
fStart%(46) = 223
fLaenge%(46) = 1
fTyp$(46) = "N"
fKonv%(46) = 1

fKurz$(47) = "Max(1)"
fBezeichnung$(47) = "Lager B:   Max. Lagerstand"
fStart%(47) = 224
fLaenge%(47) = 1
fTyp$(47) = "N"
fKonv%(47) = 1

fKurz$(48) = "Akt(1)"
fBezeichnung$(48) = "Lager B:   Akt. Lagerstand"
fStart%(48) = 225
fLaenge%(48) = 1
fTyp$(48) = "N"
fKonv%(48) = 1

fKurz$(49) = "Min(1)"
fBezeichnung$(49) = "Lager B:   Min. Lagerstand"
fStart%(49) = 226
fLaenge%(49) = 1
fTyp$(49) = "N"
fKonv%(49) = 1


fKurz$(50) = "Max(2)"
fBezeichnung$(50) = "Lager C:   Max. Lagerstand"
fStart%(50) = 227
fLaenge%(50) = 1
fTyp$(50) = "N"
fKonv%(50) = 1

fKurz$(51) = "Akt(2)"
fBezeichnung$(51) = "Lager C:   Akt. Lagerstand"
fStart%(51) = 228
fLaenge%(51) = 1
fTyp$(51) = "N"
fKonv%(51) = 1

fKurz$(52) = "Min(2)"
fBezeichnung$(52) = "Lager C:   Min. Lagerstand"
fStart%(52) = 229
fLaenge%(52) = 1
fTyp$(52) = "N"
fKonv%(52) = 1

fKurz$(53) = "Max(3)"
fBezeichnung$(53) = "Lager D:   Max. Lagerstand"
fStart%(53) = 230
fLaenge%(53) = 1
fTyp$(53) = "N"
fKonv%(53) = 1

fKurz$(54) = "Akt(3)"
fBezeichnung$(54) = "Lager D:   Akt. Lagerstand"
fStart%(54) = 231
fLaenge%(54) = 1
fTyp$(54) = "N"
fKonv%(54) = 1

fKurz$(55) = "Min(3)"
fBezeichnung$(55) = "Lager D:   Min. Lagerstand"
fStart%(55) = 232
fLaenge%(55) = 1
fTyp$(55) = "N"
fKonv%(55) = 1


fKurz$(56) = "VKDat(0)"
fBezeichnung$(56) = "VK  1 (jüngster): Datum"
fStart%(56) = 206
fLaenge%(56) = 2
fTyp$(56) = "D"
fKonv%(56) = 10

fKurz$(57) = "VKMng(0)"
fBezeichnung$(57) = "VK  1: Menge"
fStart%(57) = 208
fLaenge%(57) = 2
fTyp$(57) = "N"
fKonv%(57) = 5

fKurz$(58) = "VKDat(1)"
fBezeichnung$(58) = "VK  2: Datum"
fStart%(58) = 202
fLaenge%(58) = 2
fTyp$(58) = "D"
fKonv%(58) = 10

fKurz$(59) = "VKMng(1)"
fBezeichnung$(59) = "VK  2: Menge"
fStart%(59) = 204
fLaenge%(59) = 2
fTyp$(59) = "N"
fKonv%(59) = 5

fKurz$(60) = "VKDat(2)"
fBezeichnung$(60) = "VK  3: Datum"
fStart%(60) = 198
fLaenge%(60) = 2
fTyp$(60) = "D"
fKonv%(60) = 10

fKurz$(61) = "VKMng(2)"
fBezeichnung$(61) = "VK  2: Menge"
fStart%(61) = 200
fLaenge%(61) = 2
fTyp$(61) = "N"
fKonv%(61) = 5

fKurz$(62) = "VKDat(3)"
fBezeichnung$(62) = "VK  4: Datum"
fStart%(62) = 194
fLaenge%(62) = 2
fTyp$(62) = "D"
fKonv%(62) = 10

fKurz$(63) = "VKMng(3)"
fBezeichnung$(63) = "VK  4: Menge"
fStart%(63) = 196
fLaenge%(63) = 2
fTyp$(63) = "N"
fKonv%(63) = 5

fKurz$(64) = "VKDat(4)"
fBezeichnung$(64) = "VK  5: Datum"
fStart%(64) = 190
fLaenge%(64) = 2
fTyp$(64) = "D"
fKonv%(64) = 10

fKurz$(65) = "VKMng(4)"
fBezeichnung$(65) = "VK  5: Menge"
fStart%(65) = 192
fLaenge%(65) = 2
fTyp$(65) = "N"
fKonv%(65) = 5

fKurz$(66) = "VKDat(5)"
fBezeichnung$(66) = "VK  6: Datum"
fStart%(66) = 186
fLaenge%(66) = 2
fTyp$(66) = "D"
fKonv%(66) = 10

fKurz$(67) = "VKMng(5)"
fBezeichnung$(67) = "VK  6: Menge"
fStart%(67) = 188
fLaenge%(67) = 2
fTyp$(67) = "N"
fKonv%(67) = 5

fKurz$(68) = "VKDat(6)"
fBezeichnung$(68) = "VK  7: Datum"
fStart%(68) = 182
fLaenge%(68) = 2
fTyp$(68) = "D"
fKonv%(68) = 10

fKurz$(69) = "VKMng(6)"
fBezeichnung$(69) = "VK  7: Menge"
fStart%(69) = 184
fLaenge%(69) = 2
fTyp$(69) = "N"
fKonv%(69) = 5

fKurz$(70) = "VKDat(7)"
fBezeichnung$(70) = "VK  8: Datum"
fStart%(70) = 178
fLaenge%(70) = 2
fTyp$(70) = "D"
fKonv%(70) = 10

fKurz$(71) = "VKMng(7)"
fBezeichnung$(71) = "VK  8: Menge"
fStart%(71) = 180
fLaenge%(71) = 2
fTyp$(71) = "N"
fKonv%(71) = 5

fKurz$(72) = "VKDat(8)"
fBezeichnung$(72) = "VK  9: Datum"
fStart%(72) = 174
fLaenge%(72) = 2
fTyp$(72) = "D"
fKonv%(72) = 10

fKurz$(73) = "VKMng(8)"
fBezeichnung$(73) = "VK  9: Menge"
fStart%(73) = 176
fLaenge%(73) = 2
fTyp$(73) = "N"
fKonv%(73) = 5

fKurz$(74) = "VKDat(9)"
fBezeichnung$(74) = "VK 10 (ältester): Datum"
fStart%(74) = 170
fLaenge%(74) = 2
fTyp$(74) = "D"
fKonv%(74) = 10

fKurz$(75) = "VKMng(9)"
fBezeichnung$(75) = "VK  10: Menge"
fStart%(75) = 172
fLaenge%(75) = 2
fTyp$(75) = "N"
fKonv%(75) = 5



fKurz$(76) = "VKMon(0)"
fBezeichnung$(76) = "Monat  1 (jüngstes): Datum"
fStart%(76) = 96
fLaenge%(76) = 2
fTyp$(76) = "D"
fKonv%(76) = 10

fKurz$(77) = "VKMMng(0)"
fBezeichnung$(77) = "Monat  1: Menge"
fStart%(77) = 98
fLaenge%(77) = 2
fTyp$(77) = "N"
fKonv%(77) = 5

fKurz$(78) = "VKMon(1)"
fBezeichnung$(78) = "Monat  2: Datum"
fStart%(78) = 92
fLaenge%(78) = 2
fTyp$(78) = "D"
fKonv%(78) = 10

fKurz$(79) = "VKMMng(1)"
fBezeichnung$(79) = "Monat  2: Menge"
fStart%(79) = 94
fLaenge%(79) = 2
fTyp$(79) = "N"
fKonv%(79) = 5

fKurz$(80) = "VKMon(2)"
fBezeichnung$(80) = "Monat  3: Datum"
fStart%(80) = 88
fLaenge%(80) = 2
fTyp$(80) = "D"
fKonv%(80) = 10

fKurz$(81) = "VKMMng(2)"
fBezeichnung$(81) = "Monat  3: Menge"
fStart%(81) = 90
fLaenge%(81) = 2
fTyp$(81) = "N"
fKonv%(81) = 5

fKurz$(82) = "VKMon(3)"
fBezeichnung$(82) = "Monat  4: Datum"
fStart%(82) = 84
fLaenge%(82) = 2
fTyp$(82) = "D"
fKonv%(82) = 10

fKurz$(83) = "VKMMng(3)"
fBezeichnung$(83) = "Monat  4: Menge"
fStart%(83) = 86
fLaenge%(83) = 2
fTyp$(83) = "N"
fKonv%(83) = 5

fKurz$(84) = "VKMon(4)"
fBezeichnung$(84) = "Monat  5: Datum"
fStart%(84) = 80
fLaenge%(84) = 2
fTyp$(84) = "D"
fKonv%(84) = 10

fKurz$(85) = "VKMMng(4)"
fBezeichnung$(85) = "Monat  5: Menge"
fStart%(85) = 82
fLaenge%(85) = 2
fTyp$(85) = "N"
fKonv%(85) = 5

fKurz$(86) = "VKMon(5)"
fBezeichnung$(86) = "Monat  6: Datum"
fStart%(86) = 76
fLaenge%(86) = 2
fTyp$(86) = "D"
fKonv%(86) = 10

fKurz$(87) = "VKMMng(5)"
fBezeichnung$(87) = "Monat  6: Menge"
fStart%(87) = 78
fLaenge%(87) = 2
fTyp$(87) = "N"
fKonv%(87) = 5

fKurz$(88) = "VKMon(6)"
fBezeichnung$(88) = "Monat  7: Datum"
fStart%(88) = 72
fLaenge%(88) = 2
fTyp$(88) = "D"
fKonv%(88) = 10

fKurz$(89) = "VKMMng(6)"
fBezeichnung$(89) = "Monat  7: Menge"
fStart%(89) = 74
fLaenge%(89) = 2
fTyp$(89) = "N"
fKonv%(89) = 5

fKurz$(90) = "VKMon(7)"
fBezeichnung$(90) = "Monat  8: Datum"
fStart%(90) = 68
fLaenge%(90) = 2
fTyp$(90) = "D"
fKonv%(90) = 10

fKurz$(91) = "VKMMng(7)"
fBezeichnung$(91) = "Monat  8: Menge"
fStart%(91) = 70
fLaenge%(91) = 2
fTyp$(91) = "N"
fKonv%(91) = 5

fKurz$(92) = "VKMon(8)"
fBezeichnung$(92) = "Monat  9: Datum"
fStart%(92) = 64
fLaenge%(92) = 2
fTyp$(92) = "D"
fKonv%(92) = 10

fKurz$(93) = "VKMMng(8)"
fBezeichnung$(93) = "Monat  9: Menge"
fStart%(93) = 66
fLaenge%(93) = 2
fTyp$(93) = "N"
fKonv%(93) = 5

fKurz$(94) = "VKMon(9)"
fBezeichnung$(94) = "Monat  10: Datum"
fStart%(94) = 60
fLaenge%(94) = 2
fTyp$(94) = "D"
fKonv%(94) = 10

fKurz$(95) = "VKMMng(9)"
fBezeichnung$(95) = "Monat  10: Menge"
fStart%(95) = 62
fLaenge%(95) = 2
fTyp$(95) = "N"
fKonv%(95) = 5

fKurz$(96) = "VKMon(10)"
fBezeichnung$(96) = "Monat  11: Datum"
fStart%(96) = 56
fLaenge%(96) = 2
fTyp$(96) = "D"
fKonv%(96) = 10

fKurz$(97) = "VKMMng(10)"
fBezeichnung$(97) = "Monat  11: Menge"
fStart%(97) = 58
fLaenge%(97) = 2
fTyp$(97) = "N"
fKonv%(97) = 5

fKurz$(98) = "VKMon(11)"
fBezeichnung$(98) = "Monat  12: Datum"
fStart%(98) = 52
fLaenge%(98) = 2
fTyp$(98) = "D"
fKonv%(98) = 10

fKurz$(99) = "VKMMng(11)"
fBezeichnung$(99) = "Monat  12: Menge"
fStart%(99) = 54
fLaenge%(99) = 2
fTyp$(99) = "N"
fKonv%(99) = 5

fKurz$(100) = "VKMon(12)"
fBezeichnung$(100) = "Monat  13: Datum"
fStart%(100) = 48
fLaenge%(100) = 2
fTyp$(100) = "D"
fKonv%(100) = 10

fKurz$(101) = "VKMMng(12)"
fBezeichnung$(101) = "Monat  13: Menge"
fStart%(101) = 50
fLaenge%(101) = 2
fTyp$(101) = "N"
fKonv%(101) = 5

fKurz$(102) = "VKMon(13)"
fBezeichnung$(102) = "Monat  14: Datum"
fStart%(102) = 44
fLaenge%(102) = 2
fTyp$(102) = "D"
fKonv%(102) = 10

fKurz$(103) = "VKMMng(13)"
fBezeichnung$(103) = "Monat  14: Menge"
fStart%(103) = 46
fLaenge%(103) = 2
fTyp$(103) = "N"
fKonv%(103) = 5

fKurz$(104) = "VKMon(14)"
fBezeichnung$(104) = "Monat  15: Datum"
fStart%(104) = 40
fLaenge%(104) = 2
fTyp$(104) = "D"
fKonv%(104) = 10

fKurz$(105) = "VKMMng(14)"
fBezeichnung$(105) = "Monat  15: Menge"
fStart%(105) = 42
fLaenge%(105) = 2
fTyp$(105) = "N"
fKonv%(105) = 5

fKurz$(106) = "VKMon(15)"
fBezeichnung$(106) = "Monat  16: Datum"
fStart%(106) = 36
fLaenge%(106) = 2
fTyp$(106) = "D"
fKonv%(106) = 10

fKurz$(107) = "VKMMng(15)"
fBezeichnung$(107) = "Monat  16: Menge"
fStart%(107) = 38
fLaenge%(107) = 2
fTyp$(107) = "N"
fKonv%(107) = 5

fKurz$(108) = "VKMon(16)"
fBezeichnung$(108) = "Monat  17: Datum"
fStart%(108) = 32
fLaenge%(108) = 2
fTyp$(108) = "D"
fKonv%(108) = 10

fKurz$(109) = "VKMMng(16)"
fBezeichnung$(109) = "Monat  17: Menge"
fStart%(109) = 34
fLaenge%(109) = 2
fTyp$(109) = "N"
fKonv%(109) = 5

fKurz$(110) = "VKMon(17)"
fBezeichnung$(110) = "Monat  18: Datum"
fStart%(110) = 28
fLaenge%(110) = 2
fTyp$(110) = "D"
fKonv%(110) = 10

fKurz$(111) = "VKMMng(17)"
fBezeichnung$(111) = "Monat  18: Menge"
fStart%(111) = 30
fLaenge%(111) = 2
fTyp$(111) = "N"
fKonv%(111) = 5



fKurz$(112) = "QUART(0)"
fBezeichnung$(112) = "Verkauf Q1 (jüngstes)"
fStart%(112) = 122
fLaenge%(112) = 2
fTyp$(112) = "N"
fKonv%(112) = 5

fKurz$(113) = "QUART(1)"
fBezeichnung$(113) = "Verkauf Q2"
fStart%(113) = 124
fLaenge%(113) = 2
fTyp$(113) = "N"
fKonv%(113) = 5

fKurz$(114) = "QUART(2)"
fBezeichnung$(114) = "Verkauf Q3"
fStart%(114) = 126
fLaenge%(114) = 2
fTyp$(114) = "N"
fKonv%(114) = 5

fKurz$(115) = "QUART(3)"
fBezeichnung$(115) = "Verkauf Q4 (ältestes)"
fStart%(115) = 128
fLaenge%(115) = 2
fTyp$(115) = "N"
fKonv%(115) = 5

fKurz$(116) = "QUARTTAG"
fBezeichnung$(116) = "Tag im letzten Quartal"
fStart%(116) = 210
fLaenge%(116) = 2
fTyp$(116) = "D"
fKonv%(116) = 10

fKurz$(117) = "PK"
fBezeichnung$(117) = "Pk"
fStart%(117) = 235
fLaenge%(117) = 2
fTyp$(117) = "T"
fKonv%(117) = 0

fKurz$(118) = "LAGBEW"
fBezeichnung$(118) = "Lagerbewegung"
fStart%(118) = 212
fLaenge%(118) = 1
fTyp$(118) = "N"
fKonv%(118) = 1

fKurz$(119) = "MAX"
fBezeichnung$(119) = "Max"
fStart%(119) = 1
fLaenge%(119) = 2
fTyp$(119) = "N"
fKonv%(119) = 5



iAnzIndex% = 2
iName$(0) = "astatis.idx"
iName$(1) = "astatis2.idx"


Set Ass1 = Me

End Sub

Sub FreeClass()
Set Ass1 = Nothing
End Sub

Function OpenDatei%(Optional DateiZugriff$ = "")
Dim i%
Dim iZugriff$

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
    If (dHandle% = 0) Then
        OpenDatei% = 1
        Exit Function
    End If
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
     If (FabsErrf%) Then
        OpenDatei% = 1
        Exit Function
    End If
Next i%

OpenDatei% = 0

End Function

Sub CloseDatei()
Dim i%

If (dHandle% > 0) Then
    Close #dHandle%
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexClose(iHandle%(i%))
Next i%

End Sub

Sub SatzLock(Recno&)
Call clsDat.iLock(dHandle%, Recno&)
End Sub

Sub SatzUnLock(Recno&)
Call clsDat.iUnLock(dHandle%, Recno&)
End Sub

Sub GetRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Get #dHandle%, , assErst
Else
    Get #dHandle%, , ass
End If
End Sub

Sub GetInfoRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
Get #dHandle%, , AktData
End Sub

Function IndexSearch%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexSearch% = FabsErrf%
End Function

Function IndexGeneric%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexGeneric% = FabsErrf%
End Function

Function IndexNext%(ind%, AltRecno&, key$, Recno&)
FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexNext% = FabsErrf%
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, Recno&)
FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexPrevious% = FabsErrf%
End Function

Function IndexInsert%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexInsert% = FabsErrf%
End Function

Function SucheFeldInd%(Typ%, SuchFeld$)
Dim i%, ret%

ret% = -1

For i% = 0 To (dAnzFields% - 1)
    If (Typ%) Then
        If (fBezeichnung$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    Else
        If (fKurz$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    End If
Next i%

SucheFeldInd% = ret%

End Function

Function Info$(iFeld$)
Dim iValue As Variant
Dim ind%
Dim ret$, s$

ret$ = ""

s$ = UCase$(iFeld$)
'If (iIndex% >= 0) Then
'    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
'End If


ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    iValue = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
    ret$ = clsDat.DecodFeld(iValue, fTyp$(ind%), fKonv%(ind%))
End If

Info$ = ret$

End Function

Public Property Get FeldBezeichnung$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(0, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldBezeichnung$ = ret$

End Property

Public Property Get FeldKurz$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(1, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fKurz$(ind%)
End If

FeldKurz = ret$

End Property

Public Property Get DateiBezeichnung$()
DateiBezeichnung$ = dBezeichnung$
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property

Public Property Get DateiLen&()
DateiLen& = LOF(dHandle%)
End Property

Public Property Get wert(iFeld$, Optional iIndex% = -1) As Variant
Attribute wert.VB_UserMemId = 0
Dim ind%
Dim s$
Dim ret As Variant

ret = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
End If

wert = ret
End Property

Public Property Let wert(iFeld$, Optional iIndex% = -1, vNewValue As Variant)
Dim ind%
Dim s$, ret$

ret$ = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret$ = clsDat.MakeFeld(vNewValue, fTyp$(ind%), fKonv%(ind%))
    Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)) = Left$(ret$ + Space$(fLaenge%(ind%)), fLaenge%(ind%))
End If

End Property




Public Property Get erstmax%()
erstmax% = assErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
assErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = assErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
assErst.erstrest = vNewValue$
End Property


Public Property Get pzn$()
pzn$ = ass.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
ass.pzn = vNewValue$
End Property

Public Property Get lief() As Byte
lief = ass.lief
End Property
Public Property Let lief(ByVal vNewValue As Byte)
ass.lief = vNewValue
End Property

Public Property Get halt() As Byte
halt = ass.halt
End Property
Public Property Let halt(ByVal vNewValue As Byte)
ass.halt = vNewValue
End Property

Public Property Get lld%()
lld% = clsOpTool.CVDat(ass.lld)
End Property
Public Property Let lld(ByVal vNewValue%)
ass.lld = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get llm%()
llm% = ass.llm
End Property
Public Property Let llm(ByVal vNewValue%)
ass.llm = vNewValue%
End Property

Public Property Get abl1%()
abl1% = clsOpTool.CVDat(ass.abl1)
End Property
Public Property Let abl1(ByVal vNewValue%)
ass.abl1 = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get abl2%()
abl2% = clsOpTool.CVDat(ass.abl2)
End Property
Public Property Let abl2(ByVal vNewValue%)
ass.abl2 = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get vbm%()
vbm% = ass.vbm
End Property
Public Property Let vbm(ByVal vNewValue%)
ass.vbm = vNewValue%
End Property

Public Property Get bm%()
bm% = ass.bm
End Property
Public Property Let bm(ByVal vNewValue%)
ass.bm = vNewValue%
End Property

Public Property Get mm%()
mm% = ass.mm
End Property
Public Property Let mm(ByVal vNewValue%)
ass.mm = vNewValue%
End Property

Public Property Get lag%()
lag% = ass.lag
End Property
Public Property Let lag(ByVal vNewValue%)
ass.lag = vNewValue%
End Property

Public Property Get opt#()
opt# = ass.opt / 10#
End Property
Public Property Let opt(ByVal vNewValue#)
ass.opt = CInt(vNewValue# * 10)
End Property

Public Property Get vkmon%(ByVal ind%)
vkmon% = clsOpTool.CVDat(ass.stat(ind%).datum)
End Property
Public Property Let vkmon(ByVal ind%, ByVal vNewValue%)
ass.stat(ind%).datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get vkmmng%(ByVal ind%)
vkmmng% = ass.stat(ind%).menge
End Property
Public Property Let vkmmng(ByVal ind%, ByVal vNewValue%)
ass.stat(ind%).menge = vNewValue%
End Property

Public Property Get pp$()
pp$ = ass.pp
End Property
Public Property Let pp(ByVal vNewValue$)
ass.pp = vNewValue$
End Property

Public Property Get poslag%()
poslag% = ass.poslag
End Property
Public Property Let poslag(ByVal vNewValue%)
ass.poslag = vNewValue%
End Property

Public Property Get tplatz%()
tplatz% = ass.tplatz
End Property
Public Property Let tplatz(ByVal vNewValue%)
ass.tplatz = vNewValue%
End Property

Public Property Get tlager%()
tlager% = ass.tlager
End Property
Public Property Let tlager(ByVal vNewValue%)
ass.tlager = vNewValue%
End Property

Public Property Get tmm%()
tmm% = ass.tmm
End Property
Public Property Let tmm(ByVal vNewValue%)
ass.tmm = vNewValue%
End Property

Public Property Get EAN$()
EAN$ = ass.EAN
End Property
Public Property Let EAN(ByVal vNewValue$)
ass.EAN = vNewValue$
End Property

Public Property Get quart%(ByVal ind%)
quart% = ass.quart(ind%)
End Property
Public Property Let quart(ByVal ind%, ByVal vNewValue%)
ass.quart(ind%) = vNewValue%
End Property

Public Property Get llief%(ByVal ind%)
llief% = ass.llief(ind%).lief
End Property
Public Property Let llief(ByVal ind%, ByVal vNewValue%)
ass.llief(ind%).lief = vNewValue%
End Property

Public Property Get lldat%(ByVal ind%)
lldat% = clsOpTool.CVDat(ass.llief(ind%).datum)
End Property
Public Property Let lldat(ByVal ind%, ByVal vNewValue%)
ass.llief(ind%).datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get llrm%(ByVal ind%)
llrm% = ass.llief(ind%).lm
End Property
Public Property Let llrm(ByVal ind%, ByVal vNewValue%)
ass.llief(ind%).lm = vNewValue%
End Property

Public Property Get llnm%(ByVal ind%)
llnm% = ass.llief(ind%).nr
End Property
Public Property Let llnm(ByVal ind%, ByVal vNewValue%)
ass.llief(ind%).nr = vNewValue%
End Property

Public Property Get vkdat%(ByVal ind%)
vkdat% = clsOpTool.CVDat(ass.verk(ind%).datum)
End Property
Public Property Let vkdat(ByVal ind%, ByVal vNewValue%)
ass.verk(ind%).datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get vkmng%(ByVal ind%)
vkmng% = ass.verk(ind%).menge
End Property
Public Property Let vkmng(ByVal ind%, ByVal vNewValue%)
ass.verk(ind%).menge = vNewValue%
End Property

Public Property Get quarttag%()
quarttag% = clsOpTool.CVDat(ass.quarttag)
End Property
Public Property Let quarttag(ByVal vNewValue%)
ass.quarttag = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get lagbew() As Byte
lagbew = ass.lagbew
End Property
Public Property Let lagbew(ByVal vNewValue As Byte)
ass.lagbew = vNewValue
End Property

Public Property Get fplatz%()
fplatz% = ass.fplatz
End Property
Public Property Let fplatz(ByVal vNewValue%)
ass.fplatz = vNewValue%
End Property

Public Property Get flager%()
flager% = ass.flager
End Property
Public Property Let flager(ByVal vNewValue%)
ass.flager = vNewValue%
End Property

Public Property Get fmm%()
fmm% = ass.fmm
End Property
Public Property Let fmm(ByVal vNewValue%)
ass.fmm = vNewValue%
End Property

Public Property Get vmm%()
vmm% = ass.vmm
End Property
Public Property Let vmm(ByVal vNewValue%)
ass.vmm = vNewValue%
End Property

Public Property Get Max(ByVal ind%) As Byte
Max = ass.teinzel(ind%).teplatz
End Property
Public Property Let Max(ByVal ind%, ByVal vNewValue As Byte)
ass.teinzel(ind%).teplatz = vNewValue
End Property

Public Property Get akt(ByVal ind%) As Byte
akt = ass.teinzel(ind%).telager
End Property
Public Property Let akt(ByVal ind%, ByVal vNewValue As Byte)
ass.teinzel(ind%).telager = vNewValue
End Property

Public Property Get min(ByVal ind%) As Byte
min = ass.teinzel(ind%).temm
End Property
Public Property Let min(ByVal ind%, ByVal vNewValue As Byte)
ass.teinzel(ind%).temm = vNewValue
End Property

Public Property Get anlage%()
anlage% = clsOpTool.CVDat(ass.anlage)
End Property
Public Property Let anlage(ByVal vNewValue%)
ass.anlage = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get pk$()
pk$ = ass.pk
End Property
Public Property Let pk(ByVal vNewValue$)
ass.pk = vNewValue$
End Property

Public Property Get rest$()
rest$ = ass.rest2
End Property
Public Property Let rest(ByVal vNewValue$)
ass.rest2 = vNewValue$
End Property

Sub PrintFile(Handle%)
Put #Handle%, , ass
End Sub

