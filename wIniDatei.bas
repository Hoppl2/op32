Attribute VB_Name = "modIniDatei"
Option Explicit

Const INI_DATEI = "\user\winop.ini"

Public Const MAX_KONTROLLEN = 20
Public Const MAX_ZUORDNUNGEN = 20
Public Const MAX_RUFZEITEN = 20

Public Const KONTROLLEN_SECTION = "Kontrollen"
Public Const ZUORDNUNGEN_SECTION = "Zuordnungen"
Public Const RUFZEITEN_SECTION = "Rufzeiten"

Public Kontrollen(MAX_KONTROLLEN) As KontrollenStruct
Public Zuordnungen(MAX_ZUORDNUNGEN) As ZuordnungenStruct
Public Rufzeiten(MAX_RUFZEITEN) As RufzeitenStruct

Public AnzKontrollen%
Public AnzZuordnungen%
Public AnzRufzeiten%

Public AnzBestellWerteRows%
Public ZeigeAlleBestellZeilen%

Public AnzBestellArtikel%
Public BekartMax%
Public BekartCounter%


Public AnzeigeLaufNr&

Public BestVorsKomplett%, BestVorsKomplettMinuten%
Public BestVorsPeriodisch%, BestVorsPeriodischMinuten%

Public LiefNamen$(210)
Public AnzLiefNamen%

Public VonBM%, BisBM%
Public WaGr$, LaCo$, auto$
Public Exklusiv%

Public zTabelleAktiv%, zManuellAktiv%


Public DragX!, DragY!


Private Const DefErrModul = "winidat.bas"

Sub HoleIniKontrollen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniKontrollen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, Send$

j% = 0
For i% = 1 To MAX_KONTROLLEN
    h$ = Space$(100)
    key$ = "Kontrolle" + Format(i%, "00")
    l& = GetPrivateProfileString(KONTROLLEN_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    End If
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        op$ = RTrim$(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    End If
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert2$ = RTrim$(Left$(h$, ind% - 1))
        Send$ = Mid$(h$, ind% + 1)
        If (wert1$ <> "") Then
            Kontrollen(j%).bedingung.wert1 = wert1$
            Kontrollen(j%).bedingung.op = op$
            Kontrollen(j%).bedingung.wert2 = wert2$
            Kontrollen(j%).Send = Send$
            j% = j% + 1
        End If
    End If
Next i%
AnzKontrollen% = j%

Call DefErrPop
End Sub

Sub SpeicherIniKontrollen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniKontrollen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, h$, key$, Send$
Dim wert%, ind%, l&, lWert&

For i% = 1 To MAX_KONTROLLEN
    If (i% <= AnzKontrollen%) Then
        h$ = RTrim$(Kontrollen(i% - 1).bedingung.wert1)
        h$ = h$ + "," + RTrim$(Kontrollen(i% - 1).bedingung.op)
        h$ = h$ + "," + RTrim$(Kontrollen(i% - 1).bedingung.wert2)
        h$ = h$ + "," + RTrim$(Kontrollen(i% - 1).Send)
    Else
        h$ = ""
    End If
    
    key$ = "Kontrolle" + Format(i%, "00")
    l& = WritePrivateProfileString(KONTROLLEN_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniZuordnungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniZuordnungen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, BetrLief$, Lief2$

j% = 0
For i% = 1 To MAX_ZUORDNUNGEN
    h$ = Space$(100)
    key$ = "Zuordnung" + Format(i%, "00")
    l& = GetPrivateProfileString(ZUORDNUNGEN_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    End If
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        op$ = RTrim$(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    End If
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert2$ = RTrim$(Left$(h$, ind% - 1))
        BetrLief$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
        If (wert1$ <> "") Then
            Zuordnungen(j%).bedingung.wert1 = wert1$
            Zuordnungen(j%).bedingung.op = op$
            Zuordnungen(j%).bedingung.wert2 = wert2$
            
            For k% = 0 To 19
                ind% = InStr(BetrLief$, ",")
                If (ind% > 0) Then
                    Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                    BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                Else
                    Lief2$ = BetrLief$
                    BetrLief$ = ""
                End If
                If (Lief2$ <> "") Then
                    Zuordnungen(j%).lief(k%) = Val(Lief2$)
                End If
                If (BetrLief$ = "") Then Exit For
            Next k%
            j% = j% + 1
        End If
    End If
Next i%
AnzZuordnungen% = j%

Call DefErrPop
End Sub

Sub SpeicherIniZuordnungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniZuordnungen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, h$, key$, Send$
Dim wert%, ind%, l&, lWert&

For i% = 1 To MAX_ZUORDNUNGEN
    If (i% <= AnzZuordnungen%) Then
        h$ = RTrim$(Zuordnungen(i% - 1).bedingung.wert1)
        h$ = h$ + "," + RTrim$(Zuordnungen(i% - 1).bedingung.op)
        h$ = h$ + "," + RTrim$(Zuordnungen(i% - 1).bedingung.wert2)
        For j% = 0 To 19
            If (Zuordnungen(i% - 1).lief(j%) > 0) Then
                h$ = h$ + "," + Mid$(Str$(Zuordnungen(i% - 1).lief(j%)), 2)
            Else
                Exit For
            End If
        Next j%
    Else
        h$ = ""
    End If
    
    key$ = "Zuordnung" + Format(i%, "00")
    l& = WritePrivateProfileString(ZUORDNUNGEN_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniRufzeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniRufzeiten")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, ind%
Dim l&
Dim wert1$, RufZeit$, LieferZeit$, wert2$, h$, key$, BetrTage$, Lief2$, aErg$, aArt$, aAktiv$, aLetztSend$
Dim aGewarnt$

j% = 0
For i% = 1 To MAX_RUFZEITEN
    h$ = Space$(100)
    key$ = "Rufzeit" + Format(i%, "00")
    l& = GetPrivateProfileString(RUFZEITEN_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        Lief2$ = RTrim$(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    End If
    If (Lief2$ <> "") Then
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            RufZeit$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            LieferZeit$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            aErg$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            aArt$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            aAktiv$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            aLetztSend$ = RTrim$(Left$(h$, ind% - 1))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            aGewarnt$ = RTrim$(Left$(h$, ind% - 1))
            BetrTage$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
            
            Rufzeiten(j%).Lieferant = Val(Lief2$)
            Rufzeiten(j%).RufZeit = Val(RufZeit$)
            Rufzeiten(j%).LieferZeit = Val(LieferZeit$)
            Rufzeiten(j%).AuftragsErg = Left$(aErg$ + Space$(2), 2)
            Rufzeiten(j%).AuftragsArt = Left$(aArt$ + Space$(2), 2)
            Rufzeiten(j%).Aktiv = Left$(aAktiv$ + Space$(1), 1)
            Rufzeiten(j%).LetztSend = Val(aLetztSend$)
            Rufzeiten(j%).Gewarnt = Left$(aGewarnt$ + Space$(1), 1)
            
            For k% = 0 To 6
                ind% = InStr(BetrTage$, ",")
                If (ind% > 0) Then
                    h$ = RTrim$(Left$(BetrTage$, ind% - 1))
                    BetrTage$ = LTrim$(Mid$(BetrTage$, ind% + 1))
                Else
                    h$ = BetrTage$
                    BetrTage$ = ""
                End If
                If (h$ <> "") Then
                    Rufzeiten(j%).WoTag(k%) = Val(h$)
                End If
                If (BetrTage$ = "") Then Exit For
            Next k%
            
            j% = j% + 1
        End If
    End If
Next i%
AnzRufzeiten% = j%

Call DefErrPop
End Sub

Sub SpeicherIniRufzeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniRufzeiten")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, h$, key$, Send$
Dim wert%, ind%, l&, lWert&

For i% = 1 To MAX_RUFZEITEN
    If (i% <= AnzRufzeiten%) Then
        h$ = Mid$(Str$(Rufzeiten(i% - 1).Lieferant), 2)
        h$ = h$ + "," + Format(Rufzeiten(i% - 1).RufZeit, "0000")
        h$ = h$ + "," + Format(Rufzeiten(i% - 1).LieferZeit, "0000")
'        h$ = h$ + "," + Mid$(Str$(Rufzeiten(i% - 1).UhrZeit), 2)
        h$ = h$ + "," + RTrim$(Rufzeiten(i% - 1).AuftragsErg)
        h$ = h$ + "," + RTrim$(Rufzeiten(i% - 1).AuftragsArt)
        h$ = h$ + "," + Rufzeiten(i% - 1).Aktiv
        h$ = h$ + "," + Format(Rufzeiten(i% - 1).LetztSend, "00000000")
'        h$ = h$ + "," + Mid$(Str$(Rufzeiten(i% - 1).LetztSend), 2)
        If (Rufzeiten(i% - 1).Gewarnt = "J") Then
            h$ = h$ + "," + "J"
        Else
            h$ = h$ + "," + "N"
        End If
'        h$ = h$ + "," + Rufzeiten(i% - 1).Gewarnt
        For j% = 0 To 6
            If (Rufzeiten(i% - 1).WoTag(j%) > 0) Then
                h$ = h$ + "," + Mid$(Str$(Rufzeiten(i% - 1).WoTag(j%)), 2)
            Else
                Exit For
            End If
        Next j%
    Else
        h$ = ""
    End If
    
    key$ = "Rufzeit" + Format(i%, "00")
    l& = WritePrivateProfileString(RUFZEITEN_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Function PruefeRufzeiten%()
Dim i%, j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))

For i% = 0 To (AnzRufzeiten% - 1)
    If (Rufzeiten(i%).Lieferant > 0) Then
        For j% = 0 To 6
            rWoTag% = Rufzeiten(i%).WoTag(j%)
            rZeit% = Rufzeiten(i%).RufZeit
            rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
            If (rWoTag% > 0) Then
                If (rWoTag% = IstWoTag%) And (IstDatum& <> Rufzeiten(i%).LetztSend) Then
                    If (rZeit% = IstZeit%) And (Rufzeiten(i%).Gewarnt = "J") Then
                        Rufzeiten(i%).LetztSend = IstDatum&
                        Rufzeiten(i%).Gewarnt = "N"
                        Call SpeicherIniRufzeiten
                        PruefeRufzeiten% = i%
                        Exit Function
                    ElseIf (IstZeit% >= rZeit% - AnzMinutenWarnung%) And (IstZeit% <= rZeit%) And (Rufzeiten(i%).Gewarnt <> "J") Then
                        Rufzeiten(i%).Gewarnt = "J"
                        Call SpeicherIniRufzeiten
                        PruefeRufzeiten% = i% + 100
                        Exit Function
                    End If
                End If
            Else
                Exit For
            End If
        Next j%
    End If
Next i%

PruefeRufzeiten% = -1

End Function

Function HoleNaechstLiefernden%()
Dim i%, k%, ind%, l%, ret%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

ret% = -1

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))

With frmAction.lstSortierung
    .Clear

    For k% = 1 To 7
        For i% = 0 To (AnzRufzeiten% - 1)
            If (Rufzeiten(i%).Lieferant > 0) Then
                For j% = 0 To 6
                    rWoTag% = Rufzeiten(i%).WoTag(j%)
                    rZeit% = Rufzeiten(i%).LieferZeit
                    If (rWoTag% > 0) Then
                        If (rWoTag% = IstWoTag%) And (rZeit% > IstZeit%) Then
                            .AddItem Format(k%, "0") + Format(rZeit%, "0000") + Format(Rufzeiten(i%).Lieferant, "000")
                        End If
                    Else
                        Exit For
                    End If
                Next j%
            End If
        Next i%
    
        If (IstWoTag% = 7) Then
            IstWoTag% = 1
        Else
            IstWoTag% = IstWoTag% + 1
        End If
        IstZeit% = 0
    Next k%
    
    If (.ListCount > 0) Then
        .ListIndex = 0
        ret% = Val(Mid$(.text, 6))
    End If
End With

HoleNaechstLiefernden% = ret%

End Function

Function HoleNaechsteRufzeit%(lief%)
Dim i%, k%, ind%, l%, ret%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

ret% = -1

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))

With frmAction.lstSortierung
    .Clear

    For k% = 1 To 7
        For i% = 0 To (AnzRufzeiten% - 1)
            If (Rufzeiten(i%).Lieferant = lief%) Then
                For j% = 0 To 6
                    rWoTag% = Rufzeiten(i%).WoTag(j%)
                    rZeit% = Rufzeiten(i%).RufZeit
                    If (rWoTag% > 0) Then
                        If (rWoTag% = IstWoTag%) And (rZeit% > IstZeit%) Then
                            .AddItem Format(k%, "0") + Format(rZeit%, "0000") + Format(Rufzeiten(i%).Lieferant, "000")
                        End If
                    Else
                        Exit For
                    End If
                Next j%
            End If
        Next i%
    
        If (IstWoTag% = 7) Then
            IstWoTag% = 1
        Else
            IstWoTag% = IstWoTag% + 1
        End If
        IstZeit% = 0
    Next k%
    
    If (.ListCount > 0) Then
        .ListIndex = 0
        ret% = Val(Mid$(.text, 2, 4))
    End If
End With

HoleNaechsteRufzeit% = ret%

End Function

Sub AuslesenBestellung2()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenBestellung2")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, lief%, Max%, l%, AltLief%, ind%, row%
Dim preis#, ZeilenWert#
Dim h$, h2$, SRT$, autox$, AktPzn$
Dim lac$, BestellSort$
Static IstAktiv%

If (IstAktiv% = True) Then Call DefErrPop: Exit Sub

IstAktiv% = True

bek.SatzLock (1)
bek.GetRecord (1)

Max% = bek.erstmax
BekartCounter% = bek.erstcounter
GlobBekMax% = Max%

frmAction!lstSortierung.Clear

AnzBestellArtikel% = 0

MarkWert# = 0#
GesamtWert# = 0#

For i% = 1 To Max%
    bek.GetRecord (i% + 1)

    If (Asc(bek.pzn) < 128) Then
        preis# = bek.AEP
        ZeilenWert# = preis# * Abs(bek.bm)

        GesamtWert# = GesamtWert# + ZeilenWert#

        If (bek.aktivlief = 0) Then
' ??            If (bek.zugeordnet = "J") And ((bek.zukontrollieren = "N") Or (bek.musskontrollieren = "N") Or (bek.nochzukontrollieren = "N")) Then
            If (bek.zugeordnet = "J") And (bek.zukontrollieren <> "1") Then
                AnzBestellArtikel% = AnzBestellArtikel% + 1
                h$ = Left$(bek.txt, 18) + Mid$(bek.txt, 29) + Format(i%, "0000") + bek.pzn + Format(Abs(bek.bm), "0000")
                frmAction!lstSortierung.AddItem h$
                bek.aktivlief = Lieferant%
                bek.aktivind = AnzBestellArtikel%
                bek.PutRecord (i% + 1)
                MarkWert# = MarkWert# + ZeilenWert#
            End If
        End If
    End If
Next i%

If (AnzBestellArtikel% > 0) Then
    BekartCounter% = (BekartCounter% + 1) Mod 100
    bek.erstcounter = BekartCounter%
    bek.PutRecord (1)
End If

IstAktiv% = False

bek.SatzUnLock (1)

Call DefErrPop
End Sub

Sub ErhoeheCounter()
    
bek.GetRecord (1)
bek.erstcounter = (bek.erstcounter + 1) Mod 100
bek.PutRecord (1)
End Sub

Function CheckÄnderung%(eins$, zwei$)
'Function CheckÄnderung%(bek1 As BekartStruct, bek2 As BekartStruct)
'Dim i%, l%, ret%
'Dim temp1$, temp2$
'
'ret% = False
'
'If (bek1.best <> bek2.best) Then
'    ret% = True
'ElseIf (bek1.lief <> bek2.lief) Then
'    ret% = True
'ElseIf (bek1.bm <> bek2.bm) Then
'    ret% = True
'ElseIf (bek1.nm <> bek2.nm) Then
'    ret% = True
'ElseIf (bek1.angebot <> bek2.angebot) Then
'    ret% = True
'ElseIf (bek1.Zugeordnet <> bek2.Zugeordnet) Then
'    ret% = True
'ElseIf (bek1.ZuKontrollieren <> bek2.ZuKontrollieren) Then
'    ret% = True
'ElseIf (bek1.MussKontrollieren <> bek2.MussKontrollieren) Then
'    ret% = True
'ElseIf (bek1.NochZuKontrollieren <> bek2.NochZuKontrollieren) Then
'    ret% = True
'ElseIf (bek1.ActZuordnung <> bek2.ActZuordnung) Then
'    ret% = True
'ElseIf (bek1.BekLaufNr <> bek2.BekLaufNr) Then
'    ret% = True
'ElseIf (bek1.PosLag <> bek2.PosLag) Then
'    ret% = True
'ElseIf (bek1.AngebotInd <> bek2.AngebotInd) Then
'    ret% = True
'ElseIf (bek1.pzn <> bek2.pzn) Then
'    ret% = True
'ElseIf (bek1.txt <> bek2.txt) Then
'    ret% = True
'ElseIf (bek1.asatz <> bek2.asatz) Then
'    ret% = True
'ElseIf (bek1.ssatz <> bek2.ssatz) Then
'    ret% = True
'ElseIf (bek1.AEP <> bek2.AEP) Then
'    ret% = True
'ElseIf (bek1.abl <> bek2.abl) Then
'    ret% = True
'ElseIf (bek1.wg <> bek2.wg) Then
'    ret% = True
'ElseIf (bek1.AVP <> bek2.AVP) Then
'    ret% = True
'ElseIf (bek1.km <> bek2.km) Then
'    ret% = True
'ElseIf (bek1.absage <> bek2.absage) Then
'    ret% = True
'ElseIf (bek1.auto <> bek2.auto) Then
'    ret% = True
'ElseIf (bek1.alt <> bek2.alt) Then
'    ret% = True
'ElseIf (bek1.zr <> bek2.zr) Then
'    ret% = True
'ElseIf (bek1.besorger <> bek2.besorger) Then
'    ret% = True
'ElseIf (bek1.NNAep <> bek2.NNAep) Then
'    ret% = True
'ElseIf (bek1.nnart <> bek2.nnart) Then
'    ret% = True
'ElseIf (bek1.lief1 <> bek2.lief1) Then
'    ret% = True
'ElseIf (bek1.bm1 <> bek2.bm1) Then
'    ret% = True
'ElseIf (bek1.nm1 <> bek2.nm1) Then
'    ret% = True
'ElseIf (bek1.AbholNr <> bek2.AbholNr) Then
'    ret% = True
'Else
'    For i% = 0 To 9
'        If (bek1.ActKontrolle(i%) <> bek2.ActKontrolle(i%)) Then
'            ret% = True
'            Exit For
'        End If
'    Next i%
'End If
'
''    ActKontrolle(24) As Byte
'
''l% = Len(bek1)
''
''temp1$ = String(l%, 0)
''CopyMemory ByVal temp1$, bek1, l%
''
''temp2$ = String(l%, 0)
''CopyMemory ByVal temp2$, bek2, l%
''
''If (temp1$ <> temp2$) Then ret% = True
'
'CheckÄnderung% = ret%
End Function

Function CheckBedingung%(bed As BedingungenStruct, EK#, VK#, menge%)
Dim kGueltig%, ind%, bedTyp%
Dim val1#, val2#
Dim wert1$, wert2$, op$, h$

kGueltig% = False

wert1$ = UCase(RTrim(bed.wert1))
op$ = UCase(RTrim(bed.op))
wert2$ = UCase(RTrim(bed.wert2))
If (Mid$(wert2$, 3, 1) = ":") Then
    wert2$ = Left$(wert2$, 2) + Mid$(wert2$, 4)
End If

bedTyp% = ZeilenTyp%(wert1$)

If (bedTyp% = 0) Then

    Select Case wert1$
        Case "ABSAGEN"
            If (bek.absage > 0) Then
                kGueltig% = True
            End If
        Case "ANFRAGEN"
            If (Abs(bek.bm) = 0) And (Abs(bek.nm) = 0) Then
                kGueltig% = True
            End If
        Case "BESORGER"
            If (bek.auto = "v") Then
                kGueltig% = True
            End If
        Case "LAGERARTIKEL"
            If (bek.ssatz > 0) Then
                kGueltig% = True
            End If
        Case "DEF.LAGERARTIKEL"
            If (bek.ssatz > 0) And (bek.poslag <= 0) Then
                kGueltig% = True
            End If
        Case "MANUELLE"
            If (InStr("+v", bek.auto) = 0) Then
                kGueltig% = True
            End If
        Case "SONDERANGEBOTE"
            If (bek.angebot = "3") Or ((bek.angebot = "2") And (bek.NNAep > 0#)) Then
                kGueltig% = True
            End If
    End Select
    
    If (kGueltig%) Then
        If (op$ <> "") Then
            kGueltig% = False
            If (op$ = "VOR") Then
                If (AktUhrzeit% < Val(wert2$)) Then
                    kGueltig% = True
                End If
            ElseIf (op$ = "NACH") Then
                If (AktUhrzeit% > Val(wert2$)) Then
                    kGueltig% = True
                End If
            End If
        End If
    End If

ElseIf (bedTyp% = 1) Then

    If (wert1$ = "LIEFERANT") Then
        ind% = InStr(wert2$, "(")
        If (ind% > 0) Then
            h$ = Mid$(wert2$, ind% + 1)
            h$ = Left$(h$, Len(h$) - 1)
            If (op$ = "=") And (Val(h$) = Lieferant%) Then
                kGueltig% = True
            ElseIf (op$ = "<>") And (Val(h$) <> Lieferant%) Then
                kGueltig% = True
            End If
        End If
    ElseIf (wert1$ = "HERSTELLER") Then
        If (bek.herst = wert2$) Then
            kGueltig% = True
        End If
    End If

Else

    Select Case wert1$
        Case "BM"
            val1# = Abs(bek.bm)
            kGueltig% = True
        Case "EK"
            val1# = EK#
            kGueltig% = True
        Case "VK"
            val1# = VK#
            kGueltig% = True
        Case "ZEILENWERT"
            val1# = EK# * menge%
            kGueltig% = True
    End Select
    
    If (kGueltig%) Then
        kGueltig% = False
        
        If (wert2$ = "BMOPT") Then
            val2# = 0#
            If (bek.ssatz > 0) Then
                val2# = ass.opt
            End If
        Else
            val2# = Val(wert2$)
        End If
        
        Select Case op$
            Case "<"
                If (val1# < val2#) Then kGueltig% = True
            Case "<="
                If (val1# <= val2#) Then kGueltig% = True
            Case "="
                If (val1# = val2#) Then kGueltig% = True
            Case "<>"
                If (val1# <> val2#) Then kGueltig% = True
            Case ">="
                If (val1# >= val2#) Then kGueltig% = True
            Case ">"
                If (val1# > val2#) Then kGueltig% = True
        End Select
    End If

End If


CheckBedingung% = kGueltig%

End Function

Sub CheckKontrollen(BestellSort$, menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckKontrollen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim j%, ind%, pos%, KontrollFlag%
Dim ActKontrollen$, KontrollNr$, h$, ZuKontr$
Dim EK#, VK#

EK# = bek.AEP
VK# = bek.AVP

'ActKontrollen$ = ""
'bek.zukontrollieren = "N"
'bek.musskontrollieren = "N"
'If (bek.best = " ") Then
'    bek.nochzukontrollieren = "N"
'End If


KontrollNr$ = "0"

pos% = 0
ZuKontr$ = "4"
If (bek.aktivlief > 0) Then
    bek.zugeordnet = "Z"
ElseIf (bek.best = "B") Then
    If (bek.zukontrollieren = Chr$(0)) Then
        For j% = 0 To AnzKontrollen% - 1
            KontrollFlag% = CheckBedingung%(Kontrollen(j%).bedingung, EK#, VK#, menge%)
            
            If (KontrollFlag% = True) Then
                If (ZuKontr$ = "4") Then
                    ZuKontr$ = "2"
                End If
                If (pos% = 0) Then
                    If (Kontrollen(j%).Send = "N") Then
                        ZuKontr$ = "1"
                        KontrollNr$ = Mid$(Str$(j%), 2)
                    End If
                End If
                bek.actkontrolle(pos%) = j%
                pos% = pos% + 1
            End If
        Next j%
        If (bek.fixiert = "1") Then
            ZuKontr$ = "1"
        ElseIf (bek.fixiert = "2") Then
            ZuKontr$ = "4"
        End If
        bek.zukontrollieren = ZuKontr$
    
        For j% = pos% To 5
            bek.actkontrolle(j%) = 111
        Next j%
    Else
        ZuKontr$ = bek.zukontrollieren
    End If
Else
    For j% = pos% To 5
        bek.actkontrolle(j%) = 111
    Next j%
End If

'bek.ActKontrolle = ActKontrollen$

BestellSort$ = bek.zugeordnet + ZuKontr$ + Left$(bek.txt, 6)
    
'BestellSort$ = bek.zugeordnet + bek.zukontrollieren + bek.nochzukontrollieren + bek.musskontrollieren
'BestellSort$ = BestellSort$ + KontrollNr$ + Left$(bek.txt, 6)
    
Call DefErrPop
End Sub

Sub CheckZuordnen(lfa%, lac$, menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckZuordnen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, KontrollFlag%
Dim EK#, VK#

EK# = bek.AEP

VK# = bek.AVP
    
bek.ActZuordnung = 111

If (zManuellAktiv% And Exklusiv% And (lfa% <> Lieferant%)) Then
    bek.best = " "
    Exit Sub
End If

If (lfa% <> 0) Then
    If (lfa% = Lieferant%) Then
        bek.best = "B"
    Else
        bek.best = " "
    End If
    If (bek.best = " ") Or (zManuellAktiv% = False) Then
        Exit Sub
    End If
End If

If (zTabelleAktiv%) Then
    For j% = 0 To AnzZuordnungen% - 1
        KontrollFlag% = CheckBedingung%(Zuordnungen(j%).bedingung, EK#, VK#, menge%)
        
        If (KontrollFlag% = True) Then
            If (Zuordnungen(j%).lief(0) = 255) Then
                If (Lieferant% = NaechstLiefernderLieferant%) Then
                    bek.best = "B"
                    bek.ActZuordnung = j%
                    Exit For
                End If
            Else
'                If (Zuordnungen(j%).lief(1) = 0) Then
'                    bek.lief = Zuordnungen(j%).lief(0)
'                End If
                For i% = 0 To 19
                    If (Zuordnungen(j%).lief(i%) > 0) Then
                        If (Lieferant% = Zuordnungen(j%).lief(i%)) Then
                            bek.best = "B"
                            bek.ActZuordnung = j%
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next i%
            End If
    
            If (bek.ActZuordnung = 111) Then
                bek.best = " "
                Exit Sub
            ElseIf (zManuellAktiv% = False) Then
                Exit Sub
            End If
        End If
    Next j%
End If

If (zManuellAktiv%) Then
    If (VonBM% <> 0) And (VonBM% > menge%) Then bek.best = " ": Call DefErrPop: Exit Sub
    If (BisBM% <> 0) And (BisBM% < menge%) Then bek.best = " ": Call DefErrPop: Exit Sub
    If ((Len(WaGr$) > 0) And (InStr(WaGr$, bek.wg) = 0)) Then bek.best = " ": Call DefErrPop: Exit Sub
    If ((Len(LaCo$) > 0) And (InStr(LaCo$, lac$) = 0)) Then bek.best = " ": Call DefErrPop: Exit Sub
End If

bek.best = "B"

Call DefErrPop
End Sub

Sub ZeigeBestellZeile(Optional anzeigen% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBestellZeile")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, lief%, l%, row%, col%, iBold%, iItalic%, FirstRed%
Dim lBack&
Dim EK#
Dim h$, h2$, s$, LiefName$, ArtName$, ArtMenge$, ArtMeh$, VonWo$, nm$, zusatz$, KontrollChar$, Kz$
Dim ActKontrollen$, ActZuordnung$

KeinRowColChange% = True

LiefName$ = ""
lief% = bek.lief
'h$ = Format(lief%, "###    ")
If (lief% > 0) And (lief% < 200) Then
    lif.GetRecord (lief% + 1)
    h2$ = lif.kurz
    Call OemToChar(h2$, h2$)
    LiefName$ = h2$
End If
If (bek.fixiert = "1") Then
    LiefName$ = "(" + LiefName$ + ")"
End If

If (bek.pzn = "9999999") Then
    h$ = RTrim$(bek.txt)
    Call OemToChar(h$, h$)
    ArtName$ = h$
    ArtMenge$ = ""
    ArtMeh$ = ""
Else
    h$ = RTrim(Left$(bek.txt, 33))
    Call OemToChar(h$, h$)
    l% = Len(h$)
    For j% = l% To 2 Step -1
        h2$ = Mid$(h$, j%, 1)
        If (h2$ = " ") Then Exit For
        If (InStr(".xX", h2$) > 0) Then
            h2$ = Mid$(h$, j% - 1, 1)
            If (InStr("0123456789", h2$) <= 0) Then
                Exit For
            End If
        End If
    Next j%
    If (j% > 2) Then
        ArtName$ = Left$(h$, j%)
        ArtMenge$ = Mid$(h$, j% + 1)
    Else
        ArtName$ = h$
        ArtMenge$ = ""
    End If
    ArtMeh$ = Mid$(bek.txt, 34)
End If

EK# = bek.AEP
    
nm$ = " "
If (Abs(bek.nm) > 0) Then
    nm$ = Format(Abs(bek.nm), " 0")
End If

zusatz$ = ""
If (bek.bm = 0) And (bek.nm = 0) Then zusatz$ = "Anfrage "
If (bek.absage > 0) Then zusatz$ = zusatz$ + "Absage "

Kz$ = ""
If (bek.alt = "?") Then Kz$ = "?"
If (bek.alt = "S") Then Kz$ = Kz$ + " S"
If (bek.angebot = 2) Then
    Kz$ = Kz$ + " !"
ElseIf (bek.angebot = 3) Then
    Kz$ = Kz$ + " !"
End If
'    zusatz$ = zusatz$ + Kz$ + " "

VonWo$ = "m"
If (bek.auto = "v") Then
    VonWo$ = "v"
ElseIf (bek.auto = "+") Then
    VonWo$ = ""    '"a"
End If
'    zusatz$ = zusatz$ + VonWo$ + " "

If (Asc(Left$(bek.pzn, 1)) > 127) Then
    zusatz$ = "GELÖSCHT"
End If

ActKontrollen$ = ""
If (bek.fixiert = "1") Then
    ActKontrollen$ = "100"
End If
For j% = 0 To 5
    If (bek.actkontrolle(j%) < 111) Then
        If (ActKontrollen$ <> "") Then ActKontrollen$ = ActKontrollen$ + ","
        ActKontrollen$ = ActKontrollen$ + Mid$(Str$(bek.actkontrolle(j%)), 2)
    Else
        Exit For
    End If
Next j%

ActZuordnung$ = ""
If (bek.ActZuordnung < 111) Then
    ActZuordnung$ = Mid$(Str$(bek.ActZuordnung), 2)
End If
    
KontrollChar$ = " "
FirstRed% = False
If (bek.fixiert = "2") Then
    KontrollChar$ = Chr$(214)
ElseIf (ActKontrollen$ <> "") Then
    KontrollChar$ = "?"
    If (bek.zugeordnet = "J") Then
        FirstRed% = True
    End If
End If

'If (ActKontrollen$ <> "") Then
'    If (bek.zukontrollieren = "4") Then
'        KontrollChar$ = Chr$(214)
'    Else
'        KontrollChar$ = "?"
'        FirstRed% = True
'    End If
'End If

With frmAction.flxarbeit(0)

    col% = .col

'    If (anzeigen%) Then .redraw = False

    iItalic% = False
    lBack& = vbButtonFace

    If (Asc(Left$(bek.pzn, 1)) > 127) Or (bek.zugeordnet = "N") Or (bek.zukontrollieren = "1") Then
        lBack& = vbGrayText
    End If
    If (bek.aktivlief > 0) Then
        lBack& = vbGreen
        lif.GetRecord (bek.aktivlief + 1)
        zusatz$ = lif.kurz
        Call OemToChar(zusatz$, zusatz$)
        zusatz$ = zusatz$ + " ..."
    End If
    If (Asc(Left$(bek.pzn, 1)) > 127) Then
        iItalic% = True
    End If

    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellFontItalic = iItalic%
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle

    If (bek.angebot = 3) Then
        .col = 14
        .CellFontBold = True
    End If
    
    
    .col = 1
    .CellFontName = "Symbol"
    If (FirstRed%) Then
        .CellBackColor = vbRed
    End If
    
    
    row% = .row
    .TextMatrix(row%, 0) = bek.pzn
    .TextMatrix(row%, 1) = KontrollChar$
    .TextMatrix(row%, 2) = ArtName$
    .TextMatrix(row%, 3) = ArtMenge$
    .TextMatrix(row%, 4) = ArtMeh$
    .TextMatrix(row%, 5) = Abs(bek.bm)
    .TextMatrix(row%, 6) = nm$
    .TextMatrix(row%, 7) = LiefName$
    If (bek.ssatz = 0) Then
        .TextMatrix(row%, 8) = " "
    Else
        .TextMatrix(row%, 8) = bek.poslag
    End If
    .TextMatrix(row%, 9) = zusatz$
    .TextMatrix(row%, 10) = ActKontrollen$  'bek.ActKontrolle
    .TextMatrix(row%, 11) = Val(bek.AbholNr)
    .TextMatrix(row%, 12) = Format(EK#, "0.00")
    .TextMatrix(row%, 13) = bek.alt
    If (bek.angebot > 1) Then
        .TextMatrix(row%, 14) = "!"
    Else
        .TextMatrix(row%, 14) = " "
    End If
    .TextMatrix(row%, 15) = VonWo$
    .TextMatrix(row%, 16) = bek.lief
    .TextMatrix(row%, 17) = ActZuordnung$
    If (ProgrammChar$ = "2") Or (lBack& = vbButtonFace) Then
        .TextMatrix(row%, 19) = "*"
    Else
        .TextMatrix(row%, 19) = " "
    End If
    .TextMatrix(row%, 20) = bek.BekLaufNr
    .TextMatrix(row%, 21) = bek.aktivlief
    
    .col = col%
End With

KeinRowColChange% = False

Call DefErrPop
End Sub

Public Sub ZeigeBestellWerte(Optional AnzeigeFlag% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBestellWerte")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, AnzRows%, ind%, r%, redraw%
Dim preis#, ZeilenWert#
Dim sp&, x1&, X2&, y1&, Y2&, TextH&, TextW&, NachLinks&
Dim tx$, BestellWerte$(3, 2), BestellWerteText$(3, 2)

MarkWert# = 0#
GesamtWert# = 0#

'AnzRows% = frmAction.flxarbeit(0).Rows - 1
'For i% = 1 To AnzRows%
'    Get #BEKART%, iSRT%(i%) + 1, bek
'
'    If (Asc(Left$(bek.pzn, 1)) <= 127) Then
'        preis# = bek.AEP
'        Call DxToIEEEd(preis#)
'        ZeilenWert# = preis# * Abs(bek.bm)
'
'        GesamtWert# = GesamtWert# + ZeilenWert#
'
'        If (bek.Zugeordnet = "J") Then
'            If (bek.ZuKontrollieren = "N") Or (bek.MussKontrollieren = "N") Or (bek.NochZuKontrollieren = "N") Then
'                MarkWert# = MarkWert# + ZeilenWert#
'            End If
'        End If
'    End If
'Next i%

With frmAction.flxarbeit(0)
    AnzRows% = .Rows - 1
    For i% = 1 To AnzRows%
        tx$ = .TextMatrix(i%, 12)
        If (tx$ <> "") Then
            preis# = CDbl(.TextMatrix(i%, 12))
            ZeilenWert# = preis# * Val(.TextMatrix(i%, 5))
                
            GesamtWert# = GesamtWert# + ZeilenWert#
            If (.TextMatrix(i%, 19) = "*") Then
                MarkWert# = MarkWert# + ZeilenWert#
            End If
        End If
    Next i%
End With

If (AnzeigeFlag%) Then
    With frmAction.picBestellWerte
        .Font = frmAction.flxarbeit(0).Font
        .Cls

        sp& = .Width / 3
        x1& = sp&
        For i% = 1 To 3
            frmAction.picBestellWerte.Line (x1&, 0)-(x1&, .ScaleHeight), vbBlack
            x1& = x1& + sp&
        Next i%
    
        TextH& = .TextHeight("Äg")
        TextW& = .TextWidth("999999.99") + 2 * Screen.TwipsPerPixelX
       
        
        BestellWerteText$(0, 0) = "Gesamtwert"
        BestellWerte$(0, 0) = Format(GesamtWert#, "0.00")
        
        BestellWerteText$(0, 1) = "Ausgewählt"
        BestellWerte$(0, 1) = Format(MarkWert#, "0.00")
        
        BestellWerteText$(0, 2) = "Prozentualer Anteil"
        If (GesamtWert# > 0) Then
            tx$ = Format((MarkWert# / GesamtWert#) * 100#, "0.00")
        Else
            tx$ = Format(0#, "0.00")
        End If
        BestellWerte$(0, 2) = tx$
        
        If (AnzBestellWerteRows% > 1) Then
            Call HoleLieferantenDaten
            
            BestellWerteText$(1, 0) = "AuftragsErg"
            BestellWerte$(1, 0) = Rufzeiten(AutomaticInd%).AuftragsErg
            
            BestellWerteText$(2, 0) = "AuftragsArt"
            BestellWerte$(2, 0) = Rufzeiten(AutomaticInd%).AuftragsArt
            
            BestellWerteText$(3, 0) = "Aktivruf"
            If (Rufzeiten(AutomaticInd%).Aktiv = "J") Then
                tx$ = "ja"
            Else
                tx$ = "nein"
            End If
            BestellWerte$(3, 0) = tx$
            

            BestellWerteText$(1, 1) = "IDF-Apotheke"
            BestellWerte$(1, 1) = ApoIDF$
            
            BestellWerteText$(2, 1) = "IDF-Lieferant"
            BestellWerte$(2, 1) = GhIDF$
            
            BestellWerteText$(3, 1) = "Tel-Lieferant"
            BestellWerte$(3, 1) = TelGh$
            
            BestellWerteText$(1, 2) = "Modem"
            BestellWerteText$(2, 2) = "Parameter"
            tx$ = ZeigeModemTyp$
            ind% = InStr(tx$, "(")
            If (ind% > 0) Then
                BestellWerte$(1, 2) = Left$(tx$, ind% - 1)
                tx$ = Mid$(tx$, ind% + 1)
                BestellWerte$(2, 2) = Left$(tx$, Len(tx$) - 1)
            Else
                BestellWerte$(1, 2) = tx$
                BestellWerte$(2, 2) = ""
            End If
            
            frmAction.picBestellWerte.Line (0, 0)-(.ScaleWidth, TextH&), vbBlack, B
        End If
        
        sp& = .Width / 3
        x1& = sp& - TextW& - 4 * Screen.TwipsPerPixelX
        X2& = sp& - 2 * Screen.TwipsPerPixelX
    
        For i% = 1 To 3
            For j% = 1 To AnzBestellWerteRows%
                y1& = (j% - 1) * TextH&
                If (AnzBestellWerteRows% = 1) Then
                    Y2& = frmAction.picBestellWerte.ScaleHeight - 15
                Else
                    Y2& = j% * TextH&
                End If
                
                tx$ = BestellWerteText$(j% - 1, i% - 1)
                NachLinks& = .TextWidth(tx$)
                .CurrentX = x1& - 2 * Screen.TwipsPerPixelX - NachLinks&
                .CurrentY = y1&
                frmAction.picBestellWerte.Print tx$
                
                frmAction.picBestellWerte.Line (x1&, y1&)-(X2&, Y2&), vbWhite, BF
                frmAction.picBestellWerte.Line (x1&, y1&)-(X2&, Y2&), vbBlack, B
                
                .CurrentX = x1& + 4 * Screen.TwipsPerPixelX
                .CurrentY = y1&
                frmAction.picBestellWerte.Print BestellWerte$(j% - 1, i% - 1)
            Next j%
        
            x1& = x1& + sp&
            X2& = X2& + sp&
        Next i%
    End With
End If
        

Call DefErrPop
End Sub
   
Sub EntferneGeloeschteBestellZeilen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneGeloeschteBestellZeilen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, Max%
Dim h$, Leer$

bek.SatzLock (1)
bek.GetRecord (1)

Max% = bek.erstmax

Leer$ = String(7, 0)

i% = 1
While (i% <= Max%)
    bek.GetRecord (i% + 1)
    If ((Asc(Left$(bek.pzn, 1)) > 127) Or (bek.pzn = Leer$)) Then
        bek.GetRecord (Max% + 1)
        bek.PutRecord (i% + 1)
        Max% = Max% - 1
    Else
        i% = i% + 1
    End If
Wend

bek.erstmax = Max%
bek.PutRecord (1)
bek.SatzUnLock (1)
    
Call DefErrPop
End Sub

Function SucheFlexBestellZeile%(Optional BereitsGelockt% = False)
Dim row%, pos%, ret%, tmrStatus%
Dim LaufNr&
Dim pzn$, ch$

ret% = False

With frmAction.flxarbeit(0)
    row% = .row
    LaufNr& = Val(.TextMatrix(row%, 20))
    
    tmrStatus% = frmAction.tmrAction.Enabled
    Call frmAction.AuslesenBestellung(True, BereitsGelockt%)
    frmAction.tmrAction.Enabled = tmrStatus%
    
    row% = .row
    pzn$ = .TextMatrix(row%, 0)
    If (Val(.TextMatrix(row%, 20)) <> LaufNr&) Then
        Call MsgBox("Bestellsatz nicht mehr vorhanden!")
    Else
        pos% = Val(Right$(.TextMatrix(row%, 18), 5))
        ret% = pos%
        bek.GetRecord (pos% + 1)
        If (bek.aktivlief > 0) Then
            Call MsgBox("Bestellsatz gesperrt!")
        End If
    End If

''    ch$ = Left$(pzn$, 1)
''    If (Asc(ch$) > 127) Then
''        ch$ = Chr$(Asc(ch$) - 128)
''        Mid$(pzn$, 1, 1) = ch$
''    End If
'
'    pos% = Val(Right$(.TextMatrix(row%, 18), 5))
'
'    Get #BEKART%, pos% + 1, bek
''    If (bek.pzn = pzn$) Then
'    If (bek.BekLaufNr = LaufNr&) Then
'        SucheFlexBestellZeile% = pos%
'    Else
'        SucheFlexBestellZeile% = -1
'        Call MsgBox("DS woanders !")
'    End If
End With

SucheFlexBestellZeile% = ret%

End Function

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SatzLen% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FileOpen%")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile


If (fAttr$ = "R") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Shared As #Handle%
    Else
        Open fName$ For Random Access Read Shared As #Handle% Len = SatzLen%
    End If
    If (Err = 0) Then
        If (LOF(Handle%) = 0) Then
            Close #Handle%
            Kill (fName$)
            Err.Raise 53
        Else
            Call iLock(Handle%, 1)
            Call iUnLock(Handle%, 1)
        End If
    End If
ElseIf (fAttr$ = "W") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Write As #Handle%
    Else
        Open fName$ For Random Access Write As #Handle% Len = SatzLen%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Shared As #Handle%
    Else
        Open fName$ For Random Access Read Write Shared As #Handle% Len = SatzLen%
    End If
    Call iLock(Handle%, 1)
    Call iUnLock(Handle%, 1)
ElseIf (fAttr$ = "I") Then
    Open fName$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open fName$ For Output Access Write Shared As #Handle%
End If

If (Err = 0) Then
    FileOpen% = Handle%
Else
    Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
    Call ProgrammEnde
End If

Call DefErrPop
End Function

Function ZeilenTyp%(s$)
Dim ret%
Dim h$

h$ = UCase(Trim(s$))

ret% = 2
Select Case h$
    Case "ABSAGEN", "ANFRAGEN", "BESORGER", "DEF.LAGERARTIKEL", "MANUELLE", "LAGERARTIKEL", "SONDERANGEBOTE"
        ret% = 0
    Case "HERSTELLER", "LIEFERANT"
        ret% = 1
End Select

ZeilenTyp% = ret%

End Function

