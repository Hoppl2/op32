Attribute VB_Name = "modIniDatei"
Option Explicit

'Public Const INI_DATEI = "\user\winop.ini"  ' "\user\winop.ini"

Public Const MAX_KONTROLLEN = 50    '20
Public Const MAX_ZUORDNUNGEN = 50   '20
Public Const MAX_RUFZEITEN = 50     '20
Public Const MAX_TAETIGKEITEN = 10
Public Const MAX_AUFSCHLAEGE = 20
Public Const MAX_RUNDUNGEN = 20
Public Const MAX_WU_SORTIERUNGEN = 10
Public Const MAX_FEIERTAGE = 20
Public Const MAX_VERFALL_WARNUNGEN = 20
Public Const MAX_SIGNATUREN = 10

Public Const KALK_FREIE = 100
Public Const KALK_PREISEMPFEHLUNG = 101
Public Const KALK_AKTUELLER_PREIS = 102
Public Const KALK_AVP_EINGABE = 110

Public Const KONTROLLEN_SECTION = "Kontrollen"
Public Const ZUORDNUNGEN_SECTION = "Zuordnungen"
Public Const RUFZEITEN_SECTION = "Rufzeiten"
Public Const TAETIGKEITEN_SECTION = "Taetigkeiten"
Public Const AUFSCHLAGSTABELLE_SECTION = "Aufschlagstabelle"
Public Const RUNDUNGEN_SECTION = "Rundungen"
Public Const WU_SORTIERUNGEN_SECTION = "WU-Sortierungen"
Public Const FEIERTAGE_SECTION = "Feiertage"
Public Const VERFALL_WARNUNG_SECTION = "VerfallWarnungen"
Public Const SIGNATUREN_SECTION = "Signaturen"

Public Kontrollen(MAX_KONTROLLEN) As KontrollenStruct
Public Zuordnungen(MAX_ZUORDNUNGEN) As ZuordnungenStruct
Public Rufzeiten(MAX_RUFZEITEN) As RufzeitenStruct
Public Taetigkeiten(MAX_TAETIGKEITEN) As TaetigkeitenStruct
Public AufschlagsTabelle(MAX_AUFSCHLAEGE) As AufschlaegeStruct
Public Rundungen(MAX_RUNDUNGEN) As RundungenStruct
Public WuSortierungen(MAX_WU_SORTIERUNGEN) As WuSortierungenStruct
Public Feiertage(MAX_FEIERTAGE) As FeiertageStruct
Public VerfallWarnungen(MAX_VERFALL_WARNUNGEN) As VerfallWarnungStruct
Public Signaturen(MAX_SIGNATUREN) As SignaturenStruct

Public AnzKontrollen%
Public AnzZuordnungen%
Public AnzRufzeiten%
Public AnzTaetigkeiten%
Public AnzRundungen%
Public AnzWuSortierungen%
Public AnzFeiertage%
Public AnzVerfallWarnungen%
Public AnzSignaturen%

Public AnzBestellWerteRows%

Public AnzBestellArtikel%
Public BekartMax%
Public BekartCounter%


Public AnzeigeLaufNr&

Public BestVorsKomplett%, BestVorsKomplettMinuten%
Public BestVorsPeriodisch%, BestVorsPeriodischMinuten%

Public LiefNamen$(250)
Public AnzLiefNamen%

Public VonBM%, BisBM%
Public WaGr$, LaCo$, auto$
Public Exklusiv%

Public zTabelleAktiv%, zManuellAktiv%


Public DragX!, DragY!

Public EingabeStr$

Public GesendetDatei$

Public ActBeleg$
Public ActBelegDatum%
Public BelegModus%

Public ShuttleAktiv%

Public IstDirektLief%
Public IstAutoDirektLief%
Public DarfBestVors%

Public GhRufzeiten$, GhAllgAngebote$

Public DirektBezugSendMinuten%, DirektBezugKontrollenMinunten%, DirektBezugWarnungMinuten%
Public DirektBezugFaktRabatt#
Public DirektBezugFaktRabattTyp%, DirektBezugValutaStellung%
Public DirektBezugAbBm%

Public DirektZuordnungenHalten%

Public NachManuellerLM%

Public WuAuswahlModus%

Public Wbestk2ManuellSenden%

Public OpPartnerLiefs$
Public OpDirektPartner$

Public FremdPznOk%
Public FremdPznDB As Database
Public FremdPznRec As Recordset
Public DirektAufteilungRec As Recordset

Public OpPartnerDB As Database
Public OpPartnerRec As Recordset

Public PartnerTeilBestellungen%
Public PartnerBestaendeBeruecksichtigen%
Public KalkNichtRezPflichtigeAM%
Public LieferantenAbfrage%

Type WuStruct
    KurzEmpfänger As String * 10
    KurzAbsender As String * 10
    pzn As String * 7
    menge As String * 4
    Rmenge As String * 4
    datum As String * 6
    Verfall As String * 6
    Lief As String * 3
End Type
Public WuRec As WuStruct
Public WuSatz As String * 50

Type NnEkStruct
    KurzEmpfänger As String * 10
    KurzAbsender As String * 10
    pzn As String * 7
    BelegDatum As String * 6
    BelegNr As String * 10
    Lmenge As String * 4
    Rmenge As String * 4
    nnek As String * 8
    nnart As String * 3
    Lief As String * 3
End Type
Public NnEkRec As NnEkStruct
Public NnEkSatz As String * 65


Public WARE_HANDLE%
Public NNEK_HANDLE%
Public IdBeiPartnern$

Public AutomatenLac$
Public AutomatenLiefs$

Public SignaturBeiBuchen%

Private Const DefErrModul = "INIDATEI.BAS"

Sub HoleIniKontrollen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniKontrollen")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

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
                    Zuordnungen(j%).Lief(k%) = Val(Lief2$)
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_ZUORDNUNGEN
    If (i% <= AnzZuordnungen%) Then
        h$ = RTrim$(Zuordnungen(i% - 1).bedingung.wert1)
        h$ = h$ + "," + RTrim$(Zuordnungen(i% - 1).bedingung.op)
        h$ = h$ + "," + RTrim$(Zuordnungen(i% - 1).bedingung.wert2)
        For j% = 0 To 19
            If (Zuordnungen(i% - 1).Lief(j%) > 0) Then
                h$ = h$ + "," + Mid$(Str$(Zuordnungen(i% - 1).Lief(j%)), 2)
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
        If (ind% <= 0) And (Left$(RufZeit$, 3) = "999") Then
            h$ = h$ + ","
            ind% = InStr(h$, ",")
        End If
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

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

Sub HoleIniTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniTaetigkeiten")
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
Dim i%, j%, k%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, BetrLief$, Lief2$

j% = 0
For i% = 1 To MAX_TAETIGKEITEN
    h$ = Space$(100)
    key$ = "Taetigkeit" + Format(i%, "00")
    l& = GetPrivateProfileString(TAETIGKEITEN_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        BetrLief$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
        If (wert1$ <> "") Then
            Taetigkeiten(j%).Taetigkeit = wert1$
            
            For k% = 0 To 49
                ind% = InStr(BetrLief$, ",")
                If (ind% > 0) Then
                    Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                    BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                Else
                    Lief2$ = BetrLief$
                    BetrLief$ = ""
                End If
                If (Lief2$ <> "") Then
                    Taetigkeiten(j%).pers(k%) = Val(Lief2$)
                End If
                If (BetrLief$ = "") Then Exit For
            Next k%
            j% = j% + 1
        End If
    End If
Next i%
AnzTaetigkeiten% = j%

Call DefErrPop
End Sub

Sub SpeicherIniTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniTaetigkeiten")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_TAETIGKEITEN
    If (i% <= AnzTaetigkeiten) Then
        h$ = RTrim$(Taetigkeiten(i% - 1).Taetigkeit)
        For j% = 0 To 49
            If (Taetigkeiten(i% - 1).pers(j%) > 0) Then
                h$ = h$ + "," + Mid$(Str$(Taetigkeiten(i% - 1).pers(j%)), 2)
            Else
                Exit For
            End If
        Next j%
    Else
        h$ = ""
    End If
    
    key$ = "Taetigkeit" + Format(i%, "00")
    l& = WritePrivateProfileString(TAETIGKEITEN_SECTION$, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniSignaturen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniSignaturen")
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
Dim i%, j%, k%, ind%
Dim l&
Dim wert1$, h$, key$, Aktiv$


j% = 0
For i% = 1 To MAX_SIGNATUREN
    h$ = Space$(100)
    key$ = "Signatur" + Format(i%, "00")
    l& = GetPrivateProfileString(SIGNATUREN_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        Aktiv$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
        If (wert1$ <> "") Then
            Signaturen(j%).Taetigkeit = wert1$
            Signaturen(j%).Aktiv = Val(Aktiv$)
            j% = j% + 1
        End If
    End If
Next i%
AnzSignaturen% = j%

SignaturBeiBuchen% = 0
For i% = 0 To (AnzSignaturen% - 1)
    If (UCase(Trim(Signaturen(i%).Taetigkeit)) = "WÜ BUCHEN") Then
        SignaturBeiBuchen% = Signaturen(i%).Aktiv
        Exit For
    End If
Next i%

Call DefErrPop
End Sub

Sub HoleIniAufschlagsTabelle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniAufschlagsTabelle")
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
Dim i%, j%, k%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, BetrLief$, Lief2$

For i% = 1 To MAX_AUFSCHLAEGE
    h$ = Space$(100)
    key$ = "Aufschlag" + Format(i%, "00")
    l& = GetPrivateProfileString(AUFSCHLAGSTABELLE_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    
    h$ = Left$(h$, Len(h$) - 1)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        wert2$ = LTrim$(RTrim$(Mid$(h$, ind% + 1)))
        If (wert1$ <> "") Then
            AufschlagsTabelle(i% - 1).PreisBasis = Val(wert1$)
            AufschlagsTabelle(i% - 1).Aufschlag = Val(wert2$)
        End If
    Else
        AufschlagsTabelle(i% - 1).PreisBasis = 0
    End If
Next i%

Call DefErrPop
End Sub

Sub SpeicherIniAufschlagsTabelle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniAufschlagsTabelle")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_AUFSCHLAEGE
    h$ = ""
    If (AufschlagsTabelle(i% - 1).PreisBasis > 0) Then
        h$ = Mid$(Str$(AufschlagsTabelle(i% - 1).PreisBasis), 2)
        h$ = h$ + "," + Mid$(Str$(AufschlagsTabelle(i% - 1).Aufschlag), 2)
    End If
    
    key$ = "Aufschlag" + Format(i%, "00")
    l& = WritePrivateProfileString(AUFSCHLAGSTABELLE_SECTION$, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniRundungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniRundungen")
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
Dim i%, j%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, Gerundet$

j% = 0
For i% = 1 To MAX_RUNDUNGEN
    h$ = Space$(100)
    key$ = "Rundung" + Format(i%, "00")
    l& = GetPrivateProfileString(RUNDUNGEN_SECTION, key$, " ", h$, 101, INI_DATEI)
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
        Gerundet$ = Mid$(h$, ind% + 1)
        If (wert1$ <> "") Then
'            ind% = InStr(wert1$, "(DM)")
'            If (ind% > 0) Then
'                wert1$ = Left$(wert1$, ind% - 1) + "(EUR)" + Mid$(wert1$, ind% + 4)
'            End If
            
            Rundungen(j%).bedingung.wert1 = wert1$
            Rundungen(j%).bedingung.op = op$
            Rundungen(j%).bedingung.wert2 = wert2$
            Rundungen(j%).Gerundet = Gerundet$
            j% = j% + 1
        End If
    End If
Next i%
AnzRundungen% = j%

Call DefErrPop
End Sub

Sub SpeicherIniRundungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniRundungen")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_RUNDUNGEN
    If (i% <= AnzRundungen%) Then
        h$ = RTrim$(Rundungen(i% - 1).bedingung.wert1)
        h$ = h$ + "," + RTrim$(Rundungen(i% - 1).bedingung.op)
        h$ = h$ + "," + RTrim$(Rundungen(i% - 1).bedingung.wert2)
        h$ = h$ + "," + RTrim$(Rundungen(i% - 1).Gerundet)
    Else
        h$ = ""
    End If
    
    key$ = "Rundung" + Format(i%, "00")
    l& = WritePrivateProfileString(RUNDUNGEN_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniWuSortierungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniWuSortierungen")
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
Dim i%, j%, ind%
Dim l&
Dim wert1$, op$, wert2$, h$, key$, Send$

j% = 0
For i% = 1 To MAX_WU_SORTIERUNGEN
    h$ = Space$(100)
    key$ = "Sortierung" + Format(i%, "00")
    l& = GetPrivateProfileString(WU_SORTIERUNGEN_SECTION$, key$, " ", h$, 101, INI_DATEI)
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
        wert2$ = RTrim$(Mid$(h$, ind% + 1))
        If (wert1$ <> "") Then
            WuSortierungen(j%).bedingung.wert1 = wert1$
            WuSortierungen(j%).bedingung.op = op$
            WuSortierungen(j%).bedingung.wert2 = wert2$
            j% = j% + 1
        End If
    End If
Next i%
AnzWuSortierungen% = j%

Call DefErrPop
End Sub

Sub SpeicherIniWuSortierungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniWuSortierungen")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_WU_SORTIERUNGEN
    If (i% <= AnzWuSortierungen%) Then
        h$ = RTrim$(WuSortierungen(i% - 1).bedingung.wert1)
        h$ = h$ + "," + RTrim$(WuSortierungen(i% - 1).bedingung.op)
        h$ = h$ + "," + RTrim$(WuSortierungen(i% - 1).bedingung.wert2)
    Else
        h$ = ""
    End If
    
    key$ = "Sortierung" + Format(i%, "00")
    l& = WritePrivateProfileString(WU_SORTIERUNGEN_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniFeiertage()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniFeiertage")
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
Dim i%, j%, ind%
Dim l&
Dim wert1$, h$, key$, Aktiv$, mond$, Jahr$
Dim OsterSonntag As Date, advent1 As Date, BetTag As Date

j% = 0
For i% = 1 To MAX_FEIERTAGE
    h$ = Space$(100)
    key$ = "Feiertag" + Format(i%, "00")
    l& = GetPrivateProfileString(FEIERTAGE_SECTION$, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(Left$(h$, l&))
    If (Len(h$) <= 1) Then Exit For
    
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        Call OemToChar(wert1$, wert1$)
        Aktiv$ = Mid$(h$, ind% + 1)
        If (wert1$ <> "") Then
            Feiertage(j%).Name = wert1$
            Feiertage(j%).Aktiv = Aktiv$
            j% = j% + 1
        End If
    End If
Next i%
AnzFeiertage% = j%

If (AnzFeiertage% = 0) Then
    Feiertage(0).Name = "Neujahr"
    Feiertage(1).Name = "Heilige 3 Könige"
    Feiertage(2).Name = "Karfreitag"
    Feiertage(3).Name = "Ostermontag"
    Feiertage(4).Name = "Tag der Arbeit"
    Feiertage(5).Name = "Christi Himmelfahrt"
    Feiertage(6).Name = "Pfingstmontag"
    Feiertage(7).Name = "Fronleichnam"
    Feiertage(8).Name = "Maria Himmelfahrt"
    If (para.Land = "A") Then
        Feiertage(9).Name = "Nationalfeiertag"
    Else
        Feiertage(9).Name = "Tag der Deutschen Einheit"
    End If
    Feiertage(10).Name = "Reformationstag"
    Feiertage(11).Name = "Allerheiligen"
    If (para.Land = "A") Then
        Feiertage(12).Name = "Mariä Empfängnis"
    Else
        Feiertage(12).Name = "Buß- und Bettag"
    End If
    Feiertage(13).Name = "1. Weihnachtsfeiertag"
    Feiertage(14).Name = "2. Weihnachtsfeiertag"
    For i% = 0 To 14
        Feiertage(i%).Aktiv = "J"
    Next i%
    AnzFeiertage% = 15
End If


Jahr$ = Format(Now, "YYYY")

Select Case Val(Jahr$) Mod 19
    Case 0
        mond$ = "15.04."
    Case 1
        mond$ = "03.04."
    Case 2
        mond$ = "23.03."
    Case 3
        mond$ = "11.04."
    Case 4
        mond$ = "31.03."
    Case 5
        mond$ = "18.04."
    Case 6
        mond$ = "08.04."
    Case 7
        mond$ = "28.03."
    Case 8
        mond$ = "16.04."
    Case 9
        mond$ = "05.04."
    Case 10
        mond$ = "25.03."
    Case 11
        mond$ = "13.04."
    Case 12
        mond$ = "02.04."
    Case 13
        mond$ = "22.03."
    Case 14
        mond$ = "10.04."
    Case 15
        mond$ = "30.03."
    Case 16
        mond$ = "17.04."
    Case 17
        mond$ = "07.04."
    Case 18
        mond$ = "27.03."
End Select
mond$ = mond$ + Jahr$
  
OsterSonntag = CDate(mond$) + 7 - (WeekDay(mond$) - 1)

advent1 = "04.12." + Jahr$
While (WeekDay(advent1) <> vbSunday)
    advent1 = advent1 - 1
Wend
BetTag = advent1 - 7 - 4

Feiertage(0).KalenderTag = "01.01." + Jahr$
Feiertage(1).KalenderTag = "06.01." + Jahr$
Feiertage(2).KalenderTag = Format(OsterSonntag - 2, "DD.MM.YYYY")
Feiertage(3).KalenderTag = Format(OsterSonntag + 1, "DD.MM.YYYY")
Feiertage(4).KalenderTag = "01.05." + Jahr$
Feiertage(5).KalenderTag = Format(OsterSonntag + 39, "DD.MM.YYYY")
Feiertage(6).KalenderTag = Format(OsterSonntag + 50, "DD.MM.YYYY")
Feiertage(7).KalenderTag = Format(OsterSonntag + 60, "DD.MM.YYYY")
Feiertage(8).KalenderTag = "15.08." + Jahr$
If (para.Land = "A") Then
    Feiertage(9).KalenderTag = "26.10." + Jahr$
Else
    Feiertage(9).KalenderTag = "03.10." + Jahr$
End If
Feiertage(10).KalenderTag = "31.10." + Jahr$
Feiertage(11).KalenderTag = "01.11." + Jahr$
If (para.Land = "A") Then
    Feiertage(12).KalenderTag = "08.12." + Jahr$
Else
    Feiertage(12).KalenderTag = Format(BetTag, "DD.MM.YYYY")
End If
Feiertage(13).KalenderTag = "25.12." + Jahr$
Feiertage(14).KalenderTag = "26.12." + Jahr$

Call DefErrPop
End Sub

Sub SpeicherIniFeiertage()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniFeiertage")
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
Dim i%, j%, h$, key$, Send$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_FEIERTAGE
    If (i% <= AnzFeiertage%) Then
        h$ = RTrim$(Feiertage(i% - 1).Name)
        h$ = h$ + "," + RTrim$(Feiertage(i% - 1).Aktiv)
    Else
        h$ = ""
    End If
    Call CharToOem(h$, h$)
    
    key$ = "Feiertag" + Format(i%, "00")
    l& = WritePrivateProfileString(FEIERTAGE_SECTION$, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Sub HoleIniVerfallWarnungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniVerfallWarnungen")
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
Dim i%, j%, ind%
Dim l&
Dim wert1$, wert2$, h$, key$

j% = 0
For i% = 1 To MAX_VERFALL_WARNUNGEN
    h$ = Space$(100)
    key$ = "VerfallWarnung" + Format(i%, "00")
    l& = GetPrivateProfileString(VERFALL_WARNUNG_SECTION, key$, " ", h$, 101, INI_DATEI)
    h$ = Trim$(h$)
    If (Len(h$) <= 1) Then Exit For
    
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        wert2$ = RTrim$(Mid$(h$, ind% + 1))
        If (wert1$ <> "") Then
            VerfallWarnungen(j%).Laufzeit = Val(wert1$)
            VerfallWarnungen(j%).Warnung = Val(wert2$)
            j% = j% + 1
        End If
    End If
Next i%
AnzVerfallWarnungen% = j%

Call DefErrPop
End Sub

Sub SpeicherIniVerfallWarnungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniVerfallWarnungen")
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
Dim i%, j%, h$, key$
Dim Wert%, ind%, l&, lWert&

For i% = 1 To MAX_VERFALL_WARNUNGEN
    If (i% <= AnzVerfallWarnungen%) Then
        h$ = Trim(Str$(VerfallWarnungen(i% - 1).Laufzeit))
        If (h$ <> "") Then
            h$ = h$ + "," + Trim(Str$(VerfallWarnungen(i% - 1).Warnung))
        End If
    Else
        h$ = ""
    End If
    
    key$ = "VerfallWarnung" + Format(i%, "00")
    l& = WritePrivateProfileString(VERFALL_WARNUNG_SECTION, key$, h$, INI_DATEI)
Next i%

Call DefErrPop
End Sub

Function HoleNaechstLiefernden%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleNaechstLiefernden%")
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
Dim i%, k%, ind%, l%, ret%
Dim h$, h2$
Dim j%, IstWoTag%, lWoTag%, IstZeit%, lZeit%, rZeit%, AddTag%
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
                h$ = Format(Rufzeiten(i%).Lieferant, "000") + ","
                For j% = 0 To 6
                    lWoTag% = Rufzeiten(i%).WoTag(j%)
                    rZeit% = Rufzeiten(i%).RufZeit
                    lZeit% = Rufzeiten(i%).LieferZeit
                    If (lWoTag% > 0) Then
'                        If (lWoTag% = IstWoTag%) And (lZeit% > IstZeit%) And (rZeit% > IstZeit%) Then
'                            .AddItem Format(k%, "0") + Format(lZeit%, "0000") + Format(Rufzeiten(i%).Lieferant, "000")
'                        End If
                        If (lWoTag% = IstWoTag%) And (rZeit% >= IstZeit%) Then
                            AddTag% = 0
                            If (lZeit% < rZeit%) Then AddTag% = 1
                            .AddItem Format(k% + AddTag%, "0") + Format(lZeit%, "0000") + Format(Rufzeiten(i%).Lieferant, "000")
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

Call DefErrPop
End Function

Sub HoleRufzeitenLieferanten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleRufzeitenLieferanten")
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
Dim i%
Dim h$

GhRufzeiten$ = ""

For i% = 0 To (AnzRufzeiten% - 1)
    If (Rufzeiten(i%).Lieferant > 0) Then
        h$ = "-" + Format(Rufzeiten(i%).Lieferant, "000") + ","
        If (InStr(GhRufzeiten$, h$) <= 0) Then
            lif.GetRecord (Rufzeiten(i%).Lieferant + 1)
            GhRufzeiten$ = GhRufzeiten$ + RTrim$(lif.kurz) + h$
        End If
    End If
Next i%
    
Call DefErrPop
End Sub

Function HoleNaechsteRufzeit%(Lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleNaechsteRufzeit%")
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
            If (Rufzeiten(i%).Lieferant = Lief%) Or (Lief% = 0) Then
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
        Lief% = Val(Mid$(.text, 6))
    End If
End With

HoleNaechsteRufzeit% = ret%

Call DefErrPop
End Function

Sub HoleFruehesteManuelleSendezeit()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleFruehesteManuelleSendezeit")
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
Dim i%, k%, ind%, l%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%, ManuellSendeZeit%, iZeit%
Dim IstDatum&

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

With frmAction.lstSortierung
    .Clear

    For i% = 0 To (AnzRufzeiten% - 1)
        For j% = 0 To 6
            rWoTag% = Rufzeiten(i%).WoTag(j%)
            rZeit% = Rufzeiten(i%).RufZeit
            rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
            If (rWoTag% > 0) Then
                If (rWoTag% = IstWoTag%) Then
                    iZeit% = rZeit% + 5 'Sendedauer
                    If (Rufzeiten(i%).Aktiv <> "J") Then
                        iZeit% = iZeit% + AnzMinutenWarten%
                    End If
                    If (iZeit% > IstZeit%) Then
                        .AddItem Format(rZeit%, "0000") + Format(Rufzeiten(i%).Lieferant, "000") + Rufzeiten(i%).Aktiv
                    End If
                End If
            Else
                Exit For
            End If
        Next j%
    Next i%

    ManuellSendeZeit% = IstZeit%
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        rZeit% = Val(Mid$(.text, 1, 4))
        If ((ManuellSendeZeit% + 5) < rZeit%) Then
            Exit For
        Else
            ManuellSendeZeit% = rZeit% + 5
            If (Mid$(.text, 6, 1) <> "J") Then
                ManuellSendeZeit% = ManuellSendeZeit% + AnzMinutenWarten%
            End If
        End If
    Next i%
    
    If (ManuellSendeZeit% <> IstZeit%) Then
        h$ = Format(ManuellSendeZeit% \ 60, "00") + ":" + Format(ManuellSendeZeit% Mod 60, "00")
        Call iMsgBox("Zu erwartende früheste Sendezeit: " + h$, vbInformation)
    End If
End With

Call DefErrPop
End Sub

Function PruefeDirektBezugSendefenster%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeDirektBezugSendefenster%")
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
Dim i%, k%, ind%, l%, ret%
Dim h$, h2$
Dim j%, IstWoTag%, rWoTag%, IstZeit%, rZeit%
Dim IstDatum&

ret% = True

IstWoTag% = WeekDay(Now, vbMonday)
IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

For i% = 0 To (AnzRufzeiten% - 1)
    For j% = 0 To 6
        rWoTag% = Rufzeiten(i%).WoTag(j%)
        rZeit% = Rufzeiten(i%).RufZeit
        rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)
        If (rWoTag% > 0) Then
            If (rWoTag% = IstWoTag%) And (rZeit% > IstZeit%) And (IstZeit% + DirektBezugSendMinuten% >= rZeit%) Then
                ret% = False
            End If
        Else
            Exit For
        End If
    Next j%
Next i%

PruefeDirektBezugSendefenster% = ret%

Call DefErrPop
End Function
        
Function CheckBedingung%(Bed As BedingungenStruct, EK#, VK#, menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBedingung%")
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
Dim i%, kGueltig%, ind%, bedTyp%
Dim val1#, val2#
Dim wert1$, wert2$, op$, h$, SQLStr$

kGueltig% = False

wert1$ = UCase(RTrim(Bed.wert1))
op$ = UCase(RTrim(Bed.op))
wert2$ = UCase(RTrim(Bed.wert2))
If (Mid$(wert2$, 3, 1) = ":") Then
    wert2$ = Left$(wert2$, 2) + Mid$(wert2$, 4)
End If

bedTyp% = ZeilenTyp%(wert1$)

If (bedTyp% = 0) Then

    Select Case wert1$
        Case "ABSAGEN"
'            If (ww.absage > 0) Then
            If (ww.absage = 1) Then
                kGueltig% = True
            End If
        Case "ANFRAGEN"
            If (Abs(ww.bm) = 0) And (Abs(ww.nm) = 0) Then
                kGueltig% = True
            End If
        Case "BESORGER"
            If (ww.auto = "v") Then
                kGueltig% = True
            End If
        Case "TEXT-BESORGER"
            If (ww.pzn = "9999999") Then
                kGueltig% = True
            End If
        Case "LADENHÜTER"
            If (ww.alt = "?") Then
                kGueltig% = True
            End If
        Case "LAGERARTIKEL"
            If (ww.ssatz > 0) Then
                kGueltig% = True
            End If
        Case "DEF.LAGERARTIKEL"
            If (ww.ssatz > 0) And (ww.PosLag <= 0) Then
                kGueltig% = True
            End If
        Case "LAGERARTIKEL NEG.LS"
            If (ww.ssatz > 0) And (ww.PosLag < 0) Then
                kGueltig% = True
            End If
        Case "MANUELLE"
            If (InStr("+v", ww.auto) = 0) Then
                kGueltig% = True
            End If
        Case "AKZEPT.ANGEBOTE"
'            If (bek.angebot = "2") And (bek.nnart = 2) Then
            If (ww.angebot And &H4) Then
                kGueltig% = True
            End If
        Case "ANGEBOTSHINWEISE"
'            If (bek.angebot = "2") And (bek.nnart = 0) Then
            If (ww.angebot And &H2) And ((ww.angebot And &H4) = 0) Then
                kGueltig% = True
            End If
        Case "SCHNELLDREHER"
            If (ww.ssatz > 0) And (ww.tplatz > 0) Then
                kGueltig% = True
            End If
        Case "DEF.SCHNELLDREHER"
            If (ww.ssatz > 0) And (ww.tplatz > 0) And (ww.tlager >= ww.PosLag) Then
                kGueltig% = True
            End If
        Case "BTM"
'            If (ww.IstBtm) Then
            If (ww.ArtikelKz And KZ_BTM) Then
                kGueltig% = True
            End If
        Case "ORIGINALE", "IMPORTE"
            If (wert1$ = "ORIGINALE") And (ww.ArtikelKz And KZ_ORIGINAL) Then
                kGueltig% = True
            ElseIf (wert1$ = "IMPORTE") And (ww.ArtikelKz And KZ_IMPORT) Then
                kGueltig% = True
            End If
        Case "KÜHL/KALT"
            If (ww.ArtikelKz And KZ_KALT) Then
                kGueltig% = True
            End If
        Case "DOPPELTKONTROLLE"
            With frmAction.lstSortierung
                For i% = 0 To (.ListCount - 1)
                    If (.List(i%) = ww.pzn) Then
                        kGueltig = True
                        Exit For
                    End If
                Next i%
            End With
        Case "BEWERTUNG"
            kGueltig% = CheckBewertung%
        Case "SELBSTANGELEGTE"
            If (ww.ArtikelKz And KZ_SELBSTANGELEGT) Then
                kGueltig% = True
            End If
        Case "INTERNE STREICHUNG"
            If (ww.ssatz > 0) Then
                If (ass.halt = "S") Then
                    kGueltig% = True
                End If
            End If
        Case "AUßER HANDEL"
            If (ww.ArtikelKz And KZ_AUSSERHANDEL) Then
                kGueltig% = True
            End If
        Case "PREISG.ARTIKEL", "PREISG.ART. VORH."
            If (wert1$ = "PREISG.ARTIKEL") And (ww.ArtikelKz2 And KZ_ISTPREISGUENSTIG) Then
                kGueltig% = True
            ElseIf (wert1$ = "PREISG.ART. VORH.") And (ww.ArtikelKz2 And KZ_GIBTPREISGUENSTIG) Then
                kGueltig% = True
            End If
        Case "BESTELLUNG PARTNER"
            h$ = Format(ww.Lief, "000")
            If (InStr(OpPartnerLiefs$, h$ + ",") > 0) Then
                kGueltig% = True
            End If
        Case "UHRZEIT"
            kGueltig% = True
        Case "INTERIM"
            If (ww.ArtikelKz2 And KZ_ISTINTERIM) Then
                kGueltig% = True
            End If
        Case "REZEPTPFL.AM", "AM RX"
            If (ww.ArtikelKz2 And KZ_ISTREZPFLICHTIG) Then
                kGueltig% = True
            End If
        Case "AM NON RX"
            If (ww.asatz > 0) And ((ww.ArtikelKz2 And KZ_ISTREZPFLICHTIG) = 0) And ((ww.ArtikelKz2 And KZ_NICHTAM) = 0) Then
                kGueltig% = True
            End If
        Case "NICHTARZNEIMITTEL"
            If (ww.ArtikelKz2 And KZ_NICHTAM) Then
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
'            If (op$ = "=") And (Val(h$) = Lieferant%) Then
'                kGueltig% = True
'            ElseIf (op$ = "<>") And (Val(h$) <> Lieferant%) Then
'                kGueltig% = True
'            End If
            If (op$ = "=") And (Val(h$) = ww.Lief) Then
                kGueltig% = True
            ElseIf (op$ = "<>") And (Val(h$) <> ww.Lief) Then
                kGueltig% = True
            End If
        End If
    ElseIf (wert1$ = "HERSTELLER") Then
        If (Trim(ww.herst) = wert2$) Then
            kGueltig% = True
        End If
    ElseIf (wert1$ = "WARENGRUPPE") Then
        
        If (Len(wert2$) = 1) Then
            h$ = Left$(wert2$, 1)
            If (op$ = "=") And (h$ = ww.wg) Then
                kGueltig% = True
            ElseIf (op$ = "<>") And (h$ <> ww.wg) Then
                kGueltig% = True
            End If
        ElseIf (Len(wert2$) = 2) And (ww.asatz > 0) Then
            ast.GetRecord (ww.asatz + 1)
            If (ast.wg = wert2$) Then kGueltig% = True
        End If
        
    ElseIf (wert1$ = "LAGERCODE") Then
        If (ww.asatz > 0) Then
            ast.GetRecord (ww.asatz + 1)
            ind% = InStr(wert2$, UCase(ast.Lac))
            If (op$ = "=") And (ind% > 0) Then
                kGueltig% = True
            ElseIf (op$ = "<>") And (ind% = 0) Then
                kGueltig% = True
            End If
        End If
    
    ElseIf (wert1$ = "LAGERCODE 2ST.") Then
        If (ww.asatz > 0) Then
            If (Left$(wert2$, 1) <> ",") Then
                wert2$ = "," + wert2$
            End If
            If (Right$(wert2$, 1) <> ",") Then
                wert2$ = wert2$ + ","
            End If
            ast.GetRecord (ww.asatz + 1)
            h$ = "," + UCase(Trim(ast.Lac + ass.lac2)) + ","
            ind% = InStr(wert2$, h$)
            If (op$ = "=") And (ind% > 0) Then
                kGueltig% = True
            ElseIf (op$ = "<>") And (ind% = 0) Then
                kGueltig% = True
            End If
        End If
    End If

Else

    Select Case wert1$
        Case "BM"
            val1# = Abs(ww.bm)
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
        Case "DAUER-ABSAGEN"
            If (ww.absage > 0) Then
                val1# = ww.absage
                kGueltig% = True
            End If
        Case "ÜBL.LIEFERANT"
'            If (ww.lief > 0) Then
'                val1# = ww.lief
'                kGueltig% = True
'            End If
            If (ww.ssatz > 0) Then
                val1# = ass.Lief
                kGueltig% = True
            End If
        Case "LAGERSTAND"
            If (ww.ssatz > 0) Then
'                ass.GetRecord (ww.ssatz + 1)
'                val1# = ass.PosLag
                val1# = ww.PosLag
                kGueltig% = True
            End If
    End Select
    
    If (kGueltig%) Then
        kGueltig% = False
        
        If (wert2$ = "BMOPT") Then
            val2# = 0#
            If (ww.ssatz > 0) Then
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

Call DefErrPop
End Function

Sub CheckKontrollen(menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckKontrollen")
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
Dim j%, ind%, pos%, KontrollFlag%
Dim ActKontrollen$, KontrollNr$, h$, ZuKontr$
Dim EK#, VK#

EK# = ww.aep
VK# = ww.AVP

'ActKontrollen$ = ""
'bek.zukontrollieren = "N"
'bek.musskontrollieren = "N"
'If (bek.best = " ") Then
'    bek.nochzukontrollieren = "N"
'End If


KontrollNr$ = "0"

pos% = 0
ZuKontr$ = "4"
If (ww.aktivlief > 0) Then
    ZuKontr$ = "Z"
'    ww.zugeordnet = "Z"
'ElseIf (bek.best = "B") Then
Else
    If (ww.zukontrollieren = Chr$(0)) Then
        j% = 0
        Do
            If (j% >= AnzKontrollen%) Then Exit Do
            
            KontrollFlag% = CheckBedingung%(Kontrollen(j%).bedingung, EK#, VK#, menge%)
            
            If (KontrollFlag% = True) Then
                If (Kontrollen(j%).Send <> "U") Then
                    If (ZuKontr$ = "4") Then
                        ZuKontr$ = "2"
                    End If
                    If (pos% = 0) Then
                        If (Kontrollen(j%).Send = "N") Then
                            ZuKontr$ = "1"
                            KontrollNr$ = Mid$(Str$(j%), 2)
                        End If
                    End If
                    ww.actkontrolle(pos%) = j%
                    pos% = pos% + 1
                End If
            ElseIf (Kontrollen(j%).Send = "U") Then
                Do
                    j% = j% + 1
                    If (j% >= AnzKontrollen%) Then Exit Do
                    If (Kontrollen(j%).Send <> "U") Then Exit Do
                Loop
            End If
            
            j% = j% + 1
        Loop
        If (ww.fixiert = "1") Then
            ZuKontr$ = "1"
        ElseIf (ww.fixiert = "2") Then
            ZuKontr$ = "4"
        End If
        ww.zukontrollieren = ZuKontr$
    
        For j% = pos% To 5
            ww.actkontrolle(j%) = 111
        Next j%
    Else
        ZuKontr$ = ww.zukontrollieren
    End If
'Else
'    For j% = pos% To 5
'        bek.actkontrolle(j%) = 111
'    Next j%
End If

'bek.ActKontrolle = ActKontrollen$

If (ZuKontr$ = "4") And (ww.zugeordnet <> "J") Then ZuKontr$ = "8"
'BestellSort$ = ZuKontr$ + Left$(ww.txt, 6)

Call DefErrPop
End Sub

Sub CheckZuordnen(lfa%, Lac$, menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckZuordnen")
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
Dim i%, j%, KontrollFlag%
Dim EK#, VK#

EK# = ww.aep

VK# = ww.AVP
    
ww.actzuordnung = 111

If ((IstDirektLief%) And (lfa% <> Lieferant%) And (ww.lief1 <> Lieferant%)) Then
    ww.best = " "
    Call DefErrPop: Exit Sub
End If

If (zManuellAktiv% And Exklusiv% And (lfa% <> Lieferant%)) Then
    ww.best = " "
    Call DefErrPop: Exit Sub
End If

If (lfa% <> 0) Then
    If (lfa% = Lieferant%) Or ((IstDirektLief%) And (ww.lief1 = Lieferant%)) Then
        ww.best = "B"
    Else
        ww.best = " "
    End If
    If (ww.best = " ") Or (zManuellAktiv% = False) Then
        Call DefErrPop: Exit Sub
    End If
End If

If (zTabelleAktiv%) Then
    j% = 0
    Do
        If (j% >= AnzZuordnungen%) Then Exit Do
            
        KontrollFlag% = CheckBedingung%(Zuordnungen(j%).bedingung, EK#, VK#, menge%)
        
        If (KontrollFlag% = True) Then
            If (Zuordnungen(j%).Lief(0) < 999) Then
                If (Zuordnungen(j%).Lief(0) = 255) Then
                    If (Lieferant% = NaechstLiefernderLieferant%) Then
                        ww.best = "B"
                        ww.actzuordnung = j%
                        Exit Do
                    End If
                Else
                    For i% = 0 To 19
                        If (Zuordnungen(j%).Lief(i%) > 0) Then
                            If (Lieferant% = Zuordnungen(j%).Lief(i%)) Then
                                ww.best = "B"
                                ww.actzuordnung = j%
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next i%
                End If
        
                If (ww.actzuordnung = 111) Then
                    ww.actzuordnung = j%   'neu !!!
                    ww.best = " "
                    Call DefErrPop: Exit Sub
                ElseIf (zManuellAktiv% = False) Then
                    Call DefErrPop: Exit Sub
                End If
            End If
        ElseIf (Zuordnungen(j%).Lief(0) = 999) Then
            Do
                j% = j% + 1
                If (j% >= AnzZuordnungen%) Then Exit Do
                If (Zuordnungen(j%).Lief(0) < 999) Then Exit Do
            Loop
        End If
            
            
        j% = j% + 1
    Loop
End If

If (zManuellAktiv%) Then
    If (VonBM% <> 0) And (VonBM% > menge%) Then ww.best = " ": Call DefErrPop: Exit Sub
    If (BisBM% <> 0) And (BisBM% < menge%) Then ww.best = " ": Call DefErrPop: Exit Sub
    If ((Len(WaGr$) > 0) And (InStr(WaGr$, ww.wg) = 0)) Then ww.best = " ": Call DefErrPop: Exit Sub
    If ((Len(LaCo$) > 0) And (InStr(LaCo$, Lac$) = 0)) Then ww.best = " ": Call DefErrPop: Exit Sub
End If

ww.best = "B"

Call DefErrPop
End Sub

Function CheckWuSortierung$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckWuSortierung$")
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
Dim j%, menge%, WuSortierFlag%
Dim EK#, VK#
Dim ret$

ret$ = "Z"

EK# = ww.aep
VK# = ww.AVP
menge% = Abs(ww.bm)

For j% = 0 To AnzWuSortierungen% - 1
    WuSortierFlag% = CheckBedingung%(WuSortierungen(j%).bedingung, EK#, VK#, menge%)
    
    If (WuSortierFlag% = True) Then
        ret$ = Format(j%, "0")
        Exit For
    End If
Next j%

CheckWuSortierung$ = ret$
    
Call DefErrPop
End Function

Sub ResetZuKontrollieren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResetZuKontrollieren")
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
Dim i%, Max%
Dim h$, Leer$

ww.SatzLock (1)
ww.GetRecord (1)

Max% = ww.erstmax

Leer$ = String(7, 0)

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
'    If ((Asc(Left$(ww.pzn, 1)) <= 127) And (ww.pzn <> Leer$)) Then
    If (ww.loesch = 0) And (ww.pzn <> Leer$) Then
        ww.zukontrollieren = Chr$(0)
        ww.PutRecord (i% + 1)
    End If
Next i%

ww.SatzUnLock (1)
    
Call DefErrPop
End Sub

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SATZLEN% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FileOpen%")
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
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile


If (fAttr$ = "R") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Shared As #Handle%
    Else
        Open fName$ For Random Access Read Shared As #Handle% Len = SATZLEN%
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
        Open fName$ For Random Access Write As #Handle% Len = SATZLEN%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Shared As #Handle%
    Else
        Open fName$ For Random Access Read Write Shared As #Handle% Len = SATZLEN%
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
    Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
    Call Programmende
End If

Call DefErrPop
End Function

Sub iLock(file As Integer, SatzNr&)
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iLock")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:

If Err = 70 Or Err = 75 Then
  If LockTime = 0 Then LockTime = DateAdd("s", 20, Now)
  If LockTime > Now Then
    'Sleep (1)
    Resume
  End If
End If

Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String

If (para.MehrPlatz) Then Lock #file, SatzNr&

Call DefErrPop
End Sub

'Sub iLock(file As Integer, SatzNr&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("iLock")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim LockTime As Date
'Dim s As String
'
'On Error GoTo StandardError
'
'If (para.MehrPlatz) Then Lock #file, SatzNr&
'
'Call DefErrPop: Exit Sub
'
''------------------------------------------------------------
'StandardError:
'
'If Err = 70 Or Err = 75 Then
'  If LockTime = 0 Then LockTime = DateAdd("s", 10, Now)
'  If LockTime > Now Then
'    'Sleep (1)
'    Resume
'  End If
'End If
'Error Err.Number
'Call DefErrPop
'End Sub

Sub iUnLock(file As Integer, SatzNr&)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iUnLock")
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
If (para.MehrPlatz) Then Unlock #file, SatzNr&

Call DefErrPop
End Sub

Function ZeilenTyp%(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeilenTyp%")
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
Dim h$

h$ = UCase(Trim(s$))

'If (ProgrammChar$ = "B") Or (ProgrammChar$ = "2") Then
    ret% = 2
    Select Case h$
        Case "ABSAGEN", "ANFRAGEN", "BESORGER", "TEXT-BESORGER", "DEF.LAGERARTIKEL", "KÜHL/KALT"
            ret% = 0
        Case "MANUELLE", "LADENHÜTER", "LAGERARTIKEL", "LAGERARTIKEL NEG.LS", "AKZEPT.ANGEBOTE"
            ret% = 0
        Case "ANGEBOTSHINWEISE", "SCHNELLDREHER", "DEF.SCHNELLDREHER", "DOPPELTKONTROLLE", "BTM"
            ret% = 0
        Case "ORIGINALE", "IMPORTE", "BEWERTUNG", "SELBSTANGELEGTE", "INTERNE STREICHUNG", "AUßER HANDEL", "PREISG.ARTIKEL"
            ret% = 0
        Case "PREISG.ART. VORH.", "BESTELLUNG PARTNER", "UHRZEIT", "INTERIM", "REZEPTPFL.AM", "AM RX", "AM NON RX", "NICHTARZNEIMITTEL"
            ret% = 0
        Case "HERSTELLER", "LIEFERANT", "LAGERCODE", "LAGERCODE 2ST.", "WARENGRUPPE"
            ret% = 1
        Case "DAUER-ABSAGEN"
            ret% = 3
    End Select
'Else
    Select Case h$
        Case "1.1X"
            ret% = 0
        Case "1.XX"
            ret% = 1
        Case "ABBRUCH:TEURER(%)"
            ret% = 2
        Case "ABBRUCH:TEURER(DM)"
            ret% = 3
        Case "ABBRUCH:BILLIGER(%)"
            ret% = 4
        Case "ABBRUCH:BILLIGER(DM)"
            ret% = 5
        Case "WENN"
            ret% = 6
        Case "SONST"
            ret% = 7
    End Select
'End If

ZeilenTyp% = ret%

Call DefErrPop
End Function

Sub CheckBestzusa()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBestzusa")
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
Dim i%, bMenge%, BESTZUSA%, AnzDS%
Dim bSatz$, bPzn$, bTxt$

BESTZUSA% = FileOpen%("bestzusa.dat", "RW")
If (BESTZUSA% > 0) Then
    If (LOF(BESTZUSA%) > 0) Then
        bSatz$ = String(11, 0)
        Get #BESTZUSA%, , AnzDS%
        For i% = 1 To AnzDS%
            Seek #BESTZUSA%, i% * 11& + 1
            Get #BESTZUSA%, , bSatz$
            bPzn$ = Left$(bSatz$, 7)
            bMenge% = Val(Mid$(bSatz$, 8))
            bTxt$ = ""
            Call ManuellErfassen(bPzn$, bTxt$, False, bMenge%)
        Next i%
        
        AnzDS% = 0
        Seek #BESTZUSA%, 1
        Put #BESTZUSA%, , AnzDS%
    End If
    Close #BESTZUSA%
End If

Call DefErrPop
End Sub

Sub CheckBekart()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBekart")
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
Dim i%, j%, wwMax%, wwCounter%, bekMax%, WasGeändert%
Dim iPzn$

WasGeändert% = False

Call ww.SatzLock(1)
ww.GetRecord (1)
wwMax% = ww.erstmax
wwCounter% = ww.erstcounter

Call bek.SatzLock(1)
bek.GetRecord (1)
bekMax% = bek.erstmax

For i% = 1 To bekMax%
    bek.GetRecord (i% + 1)
        
    iPzn$ = bek.pzn
    If (Asc(Left$(iPzn$, 1)) <= 127) Then
        WasGeändert% = True
    
        ww.pzn = bek.pzn
        ww.txt = bek.txt
        ww.Lief = bek.Lief
        ww.bm = bek.bm
        ww.asatz = bek.asatz
        ww.ssatz = bek.ssatz
        ww.best = bek.best
        ww.nm = bek.nm
        ww.aep = bek.aep
        ww.abl = bek.abl
        ww.wg = bek.wg
        ww.AVP = bek.AVP
        ww.km = bek.km
        ww.absage = bek.absage
        ww.angebot = bek.angebot
        ww.auto = bek.auto
        ww.alt = bek.alt
        ww.zr = bek.zr
        ww.besorger = bek.besorger
        ww.NNAEP = bek.NNAEP
        ww.nnart = bek.nnart
        ww.lief1 = bek.lief1
        ww.bm1 = bek.bm1
        ww.nm1 = bek.nm1
        ww.AbholNr = bek.AbholNr
        
        ww.zugeordnet = Chr$(0)
        ww.zukontrollieren = Chr$(0)
        ww.fixiert = Chr$(0)
        'bek.musskontrollieren = Chr$(0)
        ww.DirektTyp = 0
        
        For j% = 0 To 5
            ww.actkontrolle(j%) = 111
        Next j%
        ww.actzuordnung = 111
        ww.aktivlief = 0
        ww.aktivind = 0
        ww.PosLag = 0
        ww.BekLaufNr = 0
        
        ww.IstSchwellArtikel = 0
        ww.OrgZeit = 0
        
        ww.loesch = 0
        ww.status = 1
        
        wwMax% = wwMax% + 1
        ww.PutRecord (wwMax% + 1)
    End If
Next i%

bek.erstmax = 0
bek.PutRecord (1)
bek.SatzUnLock (1)

If (WasGeändert%) Then
    ww.erstmax = wwMax%
    wwCounter% = (wwCounter% + 1) Mod 100
    ww.erstcounter = wwCounter%
    ww.PutRecord (1)
End If
ww.SatzUnLock (1)

Call DefErrPop
End Sub


Function KorrPzn$(Optional iPzn$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KorrPzn$")
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
Dim ch$

If (iPzn$ = "") Then
    With frmAction.flxarbeit(0)
        iPzn$ = .TextMatrix(.row, 0)
    End With
End If

ch$ = Left$(iPzn$, 1)
If (Asc(ch$) > 127) Then
    ch$ = Chr$(Asc(ch$) - 128)
    Mid$(iPzn$, 1, 1) = ch$
End If

KorrPzn$ = iPzn$

Call DefErrPop
End Function

Function KorrTxt$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KorrTxt$")
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
Dim h$

With frmAction.flxarbeit(0)
    h$ = Trim(.TextMatrix(.row, 2)) + "  " + Trim(.TextMatrix(.row, 3)) + .TextMatrix(.row, 4)
End With

KorrTxt$ = h$

Call DefErrPop
End Function

Sub ZeigeStatistik()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeStatistik")
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
Dim row%
Dim pzn$, ch$


With frmAction.flxarbeit(0)
    pzn$ = RTrim$(.TextMatrix(.row, 0))
End With

If (pzn$ = "") Then Call DefErrPop: Exit Sub

pzn$ = KorrPzn$(pzn$)

Call ZeigeStatbild(pzn$, frmAction)
AppActivate frmAction.Caption

Call DefErrPop
End Sub

Function SucheFlexZeile%(Optional BereitsGelockt% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheFlexZeile%")
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
Dim i%, row%, pos%, ret%, Max%, iBereitsGelockt%
Dim LaufNr&
Dim pzn$, ch$

ret% = False
iBereitsGelockt% = BereitsGelockt%

With frmAction.flxarbeit(0)
    row% = .row
    If (ProgrammChar$ = "B") Then
        LaufNr& = Val(.TextMatrix(row%, 21))
        pos% = Val(Right$(.TextMatrix(row%, 19), 5))
    Else
        LaufNr& = Val(.TextMatrix(row%, 20))
        pos% = Val(Right$(.TextMatrix(row%, 18), 5))
    End If
    
    If (iBereitsGelockt% = False) Then Call ww.SatzLock(1)
    ww.GetRecord (1)
    Max% = ww.erstmax
    
    ret% = SucheDateiZeile%(pos%, Max%, LaufNr&)
    
    If (ret%) Then
        If (ww.aktivlief > 0) Then
            ww.SatzUnLock (1)
            Call iMsgBox("Bestellsatz gesperrt!")
            If (iBereitsGelockt%) Then Call ww.SatzLock(1)
            iBereitsGelockt% = True
            ret% = False
        End If
    Else
        ww.SatzUnLock (1)
        Call iMsgBox("Bestellsatz nicht mehr vorhanden!")
        If (iBereitsGelockt%) Then Call ww.SatzLock(1)
        iBereitsGelockt% = True
    End If
    
    If (iBereitsGelockt% = False) Then Call ww.SatzUnLock(1)
End With

SucheFlexZeile% = ret%

Call DefErrPop
End Function

Function SucheDateiZeile%(pos%, Max%, LaufNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheDateiZeile%")
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

ret% = False

If (pos% <= Max%) Then
    ww.GetRecord (pos% + 1)
    If (ww.BekLaufNr = LaufNr&) Then
        ret% = pos%
    End If
End If

If (ret% = False) Then
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
        If (ww.BekLaufNr = LaufNr&) Then
            ret% = i%
            Exit For
        End If
    Next i%
End If

SucheDateiZeile% = ret%

Call DefErrPop
End Function

Sub EntferneGeloeschteZeilen(LoeschModus%, Optional BereitsGelockt% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneGeloeschteZeilen")
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
Dim h$, Leer$

If (BereitsGelockt% = False) Then Call ww.SatzLock(1)
ww.GetRecord (1)

Max% = ww.erstmax

Leer$ = String(7, 0)

j% = 0
For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    If ((LoeschModus% = 0) And (ww.status = 0)) Or _
       ((LoeschModus% = 1) And ((ww.loesch) Or (ww.pzn = Leer$))) Then
    Else
        j% = j% + 1
        ww.PutRecord (j% + 1)
    End If
Next i%
Max% = j%

'i% = 1
'While (i% <= Max%)
'    ww.GetRecord (i% + 1)
''    If ((Asc(Left$(ww.pzn, 1)) > 127) Or (ww.pzn = Leer$)) Then
''    If (ww.Status = 0) Or (ww.loesch) Or (ww.pzn = Leer$) Then
'    If ((LoeschModus% = 0) And (ww.Status = 0)) Or _
'       ((LoeschModus% = 1) And ((ww.loesch) Or (ww.pzn = Leer$))) Then
'        ww.GetRecord (Max% + 1)
'        ww.PutRecord (i% + 1)
'        Max% = Max% - 1
'    Else
'        i% = i% + 1
'    End If
'Wend

ww.erstmax = Max%
ww.PutRecord (1)
If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)

Call DefErrPop
End Sub

Function HoleActBenutzer%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleActBenutzer%")
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
Dim mPos1 As clsMpos

Set mPos1 = New clsMpos
mPos1.OpenDatei
mPos1.GetRecord (Val(para.User) + 1)
HoleActBenutzer% = mPos1.pwCode
mPos1.CloseDatei

Call DefErrPop
End Function


Function iMsgBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iMsgBox%")
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
Dim OrgKeinRowColChange%, ret%

OrgKeinRowColChange% = KeinRowColChange%
KeinRowColChange% = True

If (title$ <> "") Then
    ret% = MsgBox(prompt$, buttons%, title$)
Else
    ret% = MsgBox(prompt$, buttons%)
End If
KeinRowColChange% = OrgKeinRowColChange%

iMsgBox% = ret%

Call DefErrPop
End Function

Function IstFeiertag%(SuchTag$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstFeiertag%")
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

ret% = False

For i% = 0 To (AnzFeiertage% - 1)
    If (Feiertage(i%).Aktiv = "J") Then
        If (Feiertage(i%).KalenderTag = SuchTag$) Then
            ret% = True
            Exit For
        End If
    End If
Next i%

IstFeiertag% = ret%

Call DefErrPop
End Function

Sub KonvSchwellwerte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KonvSchwellwerte")
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
Dim i%, LSCHWELL%, WUMSATZ%, fehler%, tlief%
Dim aFile$, buf$
Dim buf1 As String * 8
Dim tdatum As Date, AbDatum As Date

aFile$ = "LSCHWELL.DAT"
If (Dir(aFile$) <> "") Then
    LSCHWELL% = FileOpen(aFile$, "R", "R", 8)

    For i% = 1 To lifzus.AnzRec
        lifzus.GetRecord (i% + 1)
        If (lifzus.Schwellwert(0) = 0#) Then
            Get #LSCHWELL%, i%, buf1
            lifzus.Schwellwert(0) = Val(buf1)
            lifzus.PutRecord (i% + 1)
        End If
    Next i%
    
    Close #LSCHWELL%
End If

AnzSchwellLief% = 0
ReDim SchwellLief(AnzSchwellLief%)

If (Dir$("WUMSATZ.DAT") <> "") Then
    WUMSATZ% = FileOpen("WUMSATZ.DAT", "I")
    Do While (EOF(WUMSATZ%) = False)
        Line Input #WUMSATZ%, buf$
    
        fehler% = 0
        On Error GoTo ErrorWumsatz3
        tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
        On Error GoTo DefErr
        If (fehler% = 0) Then
            tlief% = Val(Mid$(buf$, 8, 3))
            If (IstSchwellLieferant%(tlief%) < 0) Then
                ReDim Preserve SchwellLief(AnzSchwellLief%)
                SchwellLief(AnzSchwellLief%).Lief = tlief%
                AnzSchwellLief% = AnzSchwellLief% + 1
            End If
        End If
    Loop
    Close WUMSATZ%
End If

Call SpeicherSchwellwertDaten

Call DefErrPop: Exit Sub
    
ErrorWumsatz3:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

Sub EinzelSatz(typ%, Lac$, Optional MitKontrollen% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinzelSatz")
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
Dim j%, menge%, lfa%, angInd%, angLief%, erg%, DirLief%, ind%, iAngebotY%, GhBm%, GhMm%
Dim angAep#, angProzGspart#, ret#
Dim ArtikelKz As Byte, ArtikelKz2 As Byte
Dim angManuell As Byte
Dim h$, SQLStr$, iHerst$

Dim angBm%, angNm%, angBMast%, angWgAst%, angPznLagernd%
Dim angZr!, angBMopt!
Dim angAstAep#, angTaxeAep#

'If (ww.pzn = "3487126") Then
'    Beep
'End If

h$ = Trim(ww.pzn)
If (Len(h$) <> 7) Then
    ww.pzn = Right$(String(7, "0") + h$, 7)
End If

iHerst$ = Space$(Len(ww.herst))

If (ww.BekLaufNr = 0&) Then
    ww.BekLaufNr = CalcLaufNr&(ww.pzn)
End If

If (ww.OrgZeit = 0) Then ww.OrgZeit = AktUhrzeit%

menge% = Abs(ww.bm)
lfa% = ww.Lief
DirLief% = 0
GhBm% = 0

ww.ssatz = 0
FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
If (FabsErrf% = 0) Then
    ww.ssatz = CInt(FabsRecno&)
End If
If (ww.ssatz > 0) Then
    ass.GetRecord (ww.ssatz + 1)
    ww.PosLag = ass.PosLag
    ww.tplatz = ass.tplatz
    ww.tlager = ass.tlager
    
    GhMm% = ass.MM
    If ((ass.vmm > 0) Or ((ass.vmm = 0) And (ass.vbm > 0))) Then GhMm% = ass.vmm
    If (ass.PosLag <= GhMm%) Then
        GhBm% = ass.bm
        If (ass.vbm > 0) Then GhBm% = ass.vbm
    End If
End If


ww.asatz = 0
FabsErrf% = ast.IndexSearch(0, ww.pzn, FabsRecno&)
If (FabsErrf% = 0) Then
    ww.asatz = CInt(FabsRecno&)
End If
Lac$ = " "
If ((ww.asatz > 0) And ((LaCo$ <> "") Or (ww.angebot = 0))) Then
    ast.GetRecord (ww.asatz + 1)
    Lac$ = UCase(ast.Lac)
    iHerst$ = ast.herst
End If


If (ww.angebot = 0) Then

    ArtikelKz = 0
    ArtikelKz2 = 0
    h$ = Trim(ww.pzn)
    If (h$ <> "") And (h$ <> "9999999") Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + h$
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            If (TaxeRec!OriginalPZN > 0) Then
                If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                    ArtikelKz = ArtikelKz Or KZ_ORIGINAL
                Else
                    ArtikelKz = ArtikelKz Or KZ_IMPORT
                End If
            End If
            If (TaxeRec!BTMKz) Then ArtikelKz = ArtikelKz Or KZ_BTM
            If (para.Land = "A") Then
                If (InStr("ET", TaxeRec!LagerungA) > 0) Then ArtikelKz = ArtikelKz Or KZ_KALT
            Else
                If (TaxeRec!Lagerung = 1) Or (TaxeRec!Lagerung = 4) Then ArtikelKz = ArtikelKz Or KZ_KALT
            End If
            If (TaxeRec!ArtStatus = "S") Then ArtikelKz = ArtikelKz Or KZ_AUSSERHANDEL
            If (TaxeRec!AlternativPackung > 0) Then
                If (TaxeRec!AlternativPackung = TaxeRec!pzn) Then
                    ArtikelKz2 = ArtikelKz2 Or KZ_ISTPREISGUENSTIG
                Else
                    ArtikelKz2 = ArtikelKz2 Or KZ_GIBTPREISGUENSTIG
                End If
            End If
            
'            If (InStr("156", Format(TaxeRec!AbgabeBest, "0")) > 0) Then
            If (ww.asatz > 0) Then
                If (Left$(ast.rez, 1) = "+") Then
                    ArtikelKz2 = ArtikelKz2 Or KZ_ISTREZPFLICHTIG
                End If
                If (Left$(ast.rez, 1) = "P") Then
                    ArtikelKz2 = ArtikelKz2 Or KZ_NICHTAM
                End If
            End If
            
            If (ww.ArtikelKz2 And KZ_ISTINTERIM) Then
                ArtikelKz2 = ArtikelKz2 Or KZ_ISTINTERIM
            End If
            
            If (TaxeRec!Warengruppe = 18) Then
                ArtikelKz2 = ArtikelKz2 Or KZ_DOKUPFLICHT
            End If
            
            iHerst$ = TaxeRec!HerstellerKB
        Else
            ArtikelKz = ArtikelKz Or KZ_SELBSTANGELEGT
        End If
    
        FabsErrf% = ArtText.IndexSearch(0, h$, FabsRecno&)
        If (FabsErrf% = 0) Then ArtikelKz = ArtikelKz Or KZ_ZUSTEXT
        
        If (InStr(para.Benutz, "t") > 0) Then
            Set MerkzettelRec = MerkzettelDB.OpenRecordset("Merkzettel", dbOpenTable)
            MerkzettelRec.Index = "Pzn"
            MerkzettelRec.Seek "=", h$
            If (MerkzettelRec.NoMatch = False) Then ArtikelKz = ArtikelKz Or KZ_MERKZETTEL
        Else
            FabsErrf% = BESORGT.IndexSearch(0, h$, FabsRecno&)
            If (FabsErrf% = 0) Then ArtikelKz = ArtikelKz Or KZ_MERKZETTEL
        End If
    End If
    ww.ArtikelKz = ArtikelKz
    ww.ArtikelKz2 = ArtikelKz2
    
    ww.herst = iHerst$
    
    
    ww.zugeordnet = Chr$(0)
    ww.zukontrollieren = Chr$(0)
    ww.fixiert = Chr$(0)
'    ww.nochzukontrollieren = Chr$(0)   weg wegen Direktbezug Kz für Einlesen
    For j% = 0 To 5
        ww.actkontrolle(j%) = 111
    Next j%
    ww.actzuordnung = 111
    ww.aktivlief = 0
    ww.aktivind = 0

    '2.69 alte Werte besser vor der Verõnderung speichern
    ww.bm1 = ww.bm
    ww.lief1 = ww.Lief
    ww.nm1 = ww.nm

'    angLief% = lfa%
'    If (lfa% > 0) And (ww.auto = "+") And (ww.bm <> 0) Then
'        lifzus.GetRecord (lfa% + 1)
'        If (lifzus.IstDirektLieferant) Then
'            angLief% = 0
'            iHerst$ = iHerst$ + vbCr + Format(lfa%, "000")
'        End If
'    End If
    
    angLief% = lfa%
    If (lfa% > 0) Then
        lifzus.GetRecord (lfa% + 1)
        If (lifzus.IstDirektLieferant) Then
            If (lifzus.ZuordnungenAktiv) Or (DirektZuordnungenHalten%) Then
                If (ww.auto = "+") And (ww.bm <> 0) Then
                    DirLief% = lfa%
                    angLief% = 0
                    iHerst$ = iHerst$ + vbCr + Format(lfa%, "000")
                End If
            Else
                lfa% = 0
                ww.Lief = 0
                ww.lief1 = 0
                angLief% = 0
            End If
        End If
    End If
    
    
    If (ArtikelKz2 And KZ_ISTINTERIM) Then
        erg% = 0
        angInd% = 1
    Else
        angBm% = Abs(ww.bm)
        erg% = ErstAngebot(ww.pzn, angInd%, angManuell, angBm%, angNm%, angLief%, angAep#, angZr!, iHerst$, GhAllgAngebote$)
    End If
    
    ww.angebot = angInd%
    
    If (erg%) Then
        If (ww.auto = "+") Then                        '2.69
            If (AutomatenLiefs$ <> "") And (ww.asatz > 0) Then
                If (AutomatenLac$ = Lac$) Then
                    h$ = "," + Format(angLief%, "000") + "-"
                    ind% = InStr(AutomatenLiefs$, h$)
                    If (ind% > 0) Then
                        angLief% = Val(Mid$(AutomatenLiefs$, ind% + 5, 3))
                    End If
                End If
            End If
'            If (DirLief% > 0) And (angLief% <> DirLief%) Then
'                If (GhBm% > 0) Then
'                    ww.bm = GhBm%
'                    ww.nm = 0
'                    ww.Lief = angLief%
'                Else
'                    ww.bm = 0
'                    ww.nm = 0
'                    ww.Lief = DirLief%
'                End If
'                ww.best = " "
'            Else
                ww.angebot = ww.angebot Or &H4
                If (angManuell) Then ww.angebot = ww.angebot Or &H20
                ww.AngebotInd = erg%    ' 0
                If (angAep# > 0#) Then
                    ww.NNAEP = FNX(angAep#)
                    ww.nnart = 2
                End If
                ww.bm = angBm%
                ww.nm = angNm%
                ww.Lief = angLief%
                ww.zr = angZr!
                ww.best = " "
'            End If
        End If    '2.56b->
    ElseIf (FremdPznOk%) And (ww.auto <> "v") And (DirLief% = 0) Then
        Call PruefePartnerLadenhueter
    End If

End If

menge% = Abs(ww.bm)
lfa% = ww.Lief




If (typ% = 0) Then
    Call CheckZuordnen(lfa%, Lac$, menge%)
    
    If (ww.best = "B") Then
        ww.zugeordnet = "J"
    Else
        ww.zugeordnet = "N"
    End If
End If

If (MitKontrollen%) Then Call CheckKontrollen(menge%)
    
Call DefErrPop
End Sub

Function HoleIniString$(key$, Optional Section$ = "Bestellung")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniString$")
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
Dim l&
Dim h$

h$ = Space$(255)
l& = GetPrivateProfileString(Section$, key$, h$, h$, 255, INI_DATEI)
HoleIniString$ = Left$(h$, l&)
   
Call DefErrPop
End Function

Function SpeicherIniString%(key$, Value$, Optional Section$ = "Bestellung")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniString%")
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
Dim l&

l& = WritePrivateProfileString(Section$, key$, Value$, INI_DATEI)

SpeicherIniString% = True

Call DefErrPop
End Function

Function CheckBewertung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBewertung%")
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
Dim ret%, iLief%, HatBewertung%, IstBewertungOk%, AngebotY%
Dim angBm%, angNm%
Dim angZr!
Dim angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%, dRet#
Dim SQLStr$


'Dim rInd%, angInd%, angBm%, angNm%, angLief%, angAep#, angIndOrg%, iAngebot%, angAngebotYOrg%
'Dim iiLief%, iaLief%
'Dim l&
'Dim h$, ch$


'Dim i%, oldRow%, oldCol%, oldStatus%, AnzRows%, spBreite%, Y%
'Dim sp&
'Dim ret#
'Dim tx$, FormStr$


ret% = False

HatBewertung% = False
iLief% = ww.Lief
If (iLief% > 0) Then
    lifzus.GetRecord (iLief% + 1)
    HatBewertung% = lifzus.IstDirektLieferant
End If

If (HatBewertung%) And (ww.aktivlief = 0) And (Abs(ww.bm) > 0) Then

    IstBewertungOk% = 3

    Call InitDirektBewertung(iLief%)
    
    angBm% = Abs(ww.bm)
    angNm% = ww.nm
    angZr! = ww.zr
    
    
    angBMopt! = 0
    angBMast% = 0
    angAstAep# = 0
    angTaxeAep# = 0
    angWgAst% = 0
    angPznLagernd% = 0
    
    If (ww.ssatz > 0) Then
        angBMopt! = ass.opt
        angBMast% = ass.bm
    End If
    
    If (ww.asatz > 0) Then
        ast.GetRecord (ww.asatz + 1)
        angAstAep# = ast.aep
    End If
    
    
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
        angTaxeAep# = TaxeRec!EK / 100
        If (ww.asatz <= 0) Then
            Call Taxe2ast(ww.pzn)
        End If
        angWgAst% = ast.wg
    End If
    
    If (ww.asatz > 0) Or (TaxeRec.EOF = False) Then angPznLagernd% = 1
    
    Call RechneDirektBewertung(ww.pzn, angBm%, angNm%, angZr!, angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%)

    dRet# = ZeigeDirektBewertung(IstBewertungOk%, AngebotY%, False)
'    If (IstBewertungOk% = False) Then ret% = True
    If (IstBewertungOk% > 0) Then ret% = True
End If

CheckBewertung% = ret%

Call DefErrPop
End Function

Function SucheRueckKaufZeile%(pos%, Max%, LaufNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheRueckKaufZeile%")
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

ret% = False

If (pos% <= Max%) Then
    rk.GetRecord (pos% + 1)
    If (rk.BekLaufNr = LaufNr&) Then
        ret% = pos%
    End If
End If

If (ret% = False) Then
    For i% = 1 To Max%
        rk.GetRecord (i% + 1)
        If (rk.BekLaufNr = LaufNr&) Then
            ret% = i%
            Exit For
        End If
    Next i%
End If

SucheRueckKaufZeile% = ret%

Call DefErrPop
End Function

Sub EntferneGeloeschteRueckKauf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneGeloeschteRueckKauf")
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

rk.GetRecord (1)

Max% = rk.erstmax

j% = 0
For i% = 1 To Max%
    rk.GetRecord (i% + 1)
    If (rk.status > 0) And (rk.loesch = 0) Then
        j% = j% + 1
        rk.PutRecord (j% + 1)
    End If
Next i%
Max% = j%

rk.erstmax = Max%
rk.PutRecord (1)

Call DefErrPop
End Sub

Sub PruefePartnerLadenhueter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefePartnerLadenhueter")
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
Dim iProfilNr%, iMaxLiefNr%, iMaxPOS%, ok%
Dim pzn$

ok% = 0
iMaxLiefNr% = -1
iMaxPOS% = -999
pzn$ = ww.pzn

FremdPznRec.Seek "=", pzn$
If (FremdPznRec.NoMatch = False) Then
    Do
        If (FremdPznRec.EOF) Then
            Exit Do
        End If
        If (FremdPznRec!pzn <> pzn$) Then
            Exit Do
        End If
        
        If (FremdPznRec!Ladenhüter) Then
            iProfilNr% = FremdPznRec!ProfilNr
            If (iProfilNr% > 0) Then
                OpPartnerRec.Seek "=", iProfilNr%
                If (OpPartnerRec.NoMatch = False) Then
                    If (OpPartnerRec!BeiGh) Then
                        If (PartnerBestaendeBeruecksichtigen% = 0) Or (FremdPznRec!pos >= Abs(ww.bm)) Then
                            ww.Lief = OpPartnerRec!IntLiefNr
                            Call ErhoehePartnerCounter
                            ok% = True
                            Exit Do
                        ElseIf (PartnerTeilBestellungen%) Then
                            If (FremdPznRec!pos > iMaxPOS%) Then
                                iMaxLiefNr% = OpPartnerRec!IntLiefNr
                                iMaxPOS% = FremdPznRec!pos
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        FremdPznRec.MoveNext
    Loop
    
    If (PartnerTeilBestellungen%) And (ok% = 0) And (iMaxPOS% > 0) Then
        ww.Lief = iMaxLiefNr%
        Call ErhoehePartnerCounter
        ww.bm = iMaxPOS%
    End If
End If
        
Call DefErrPop
End Sub

Sub EinlesenOpPartnerLiefs()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenOpPartnerLiefs")
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
Dim h$, h2$

Set OpPartnerRec = OpPartnerDB.OpenRecordset("PartnerProfile", dbOpenTable)
OpPartnerRec.Index = "Unique"
If (OpPartnerRec.RecordCount > 0) Then
    OpPartnerRec.MoveFirst
End If

h$ = ""
h2$ = ""
Do
    If (OpPartnerRec.EOF) Then Exit Do
    
    h$ = h$ + Format(OpPartnerRec!IntLiefNr, "000") + ","
    
    If (OpPartnerRec!BeiDirektbezug) Then
        h2$ = h2$ + Format(OpPartnerRec!ProfilNr, "000") + ","
    End If
    
    OpPartnerRec.MoveNext
Loop

OpPartnerLiefs$ = h$
OpDirektPartner$ = h2$

Call DefErrPop
End Sub

Sub ErhoehePartnerCounter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErhoehePartnerCounter")
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
Dim iCou%
Dim l&
Dim h$

h$ = "000"
l& = GetPrivateProfileString("Bestellung", "PartnerCounter", h$, h$, 4, INI_DATEI)
h$ = Left$(h$, l&)
iCou% = Val(h$) + 1
If (iCou% >= 100) Then
    iCou% = 1
End If
l& = WritePrivateProfileString("Bestellung", "PartnerCounter", Str$(iCou%), INI_DATEI)

Call DefErrPop
End Sub

Function OpenCreateMerkzettelDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenCreateMerkzettelDB%")
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
Dim ret%, DbNeu%, Max%, OpenErg%, ErrNumber%
Dim i&, DateiMax&
Dim DBname$
Dim Td As TableDef
Dim Idx As Index
Dim Fld As Field
Dim IxFld As Field

ret% = True
DbNeu% = False

If (InStr(para.Benutz, "t") > 0) Then
    DBname$ = Merkzettel.DateiName
    On Error Resume Next
    Err.Clear
    Set MerkzettelDB = OpenDatabase(DBname$, False, False)
    ErrNumber% = Err.Number
    On Error GoTo DefErr
    If (ErrNumber% > 0) Then
        DbNeu% = True
    
        If Dir(DBname$) <> "" Then Kill DBname$
        Set MerkzettelDB = CreateDatabase(DBname$, dbLangGeneral)
    
    'Tabelle Merkzettel
        Set Td = MerkzettelDB.CreateTableDef("Merkzettel")
    
        Set Fld = Td.CreateField("Pzn", dbText)
        Fld.AllowZeroLength = False
        Fld.Size = 7
        Td.Fields.Append Fld
    
        Set Fld = Td.CreateField("Txt", dbText)
        Fld.AllowZeroLength = False
        Fld.Size = 36
        Td.Fields.Append Fld
    
        Set Fld = Td.CreateField("BestellDatum", dbDate)
        Td.Fields.Append Fld
        
        Set Fld = Td.CreateField("LieferDatum", dbDate)
        Td.Fields.Append Fld
        
        Set Fld = Td.CreateField("Lief", dbInteger)
        Td.Fields.Append Fld
    
        Set Fld = Td.CreateField("Lm", dbInteger)
        Td.Fields.Append Fld
    
        Set Fld = Td.CreateField("Loesch", dbByte)
        Td.Fields.Append Fld
    
        ' Indizes für Merkzettel
        Set Idx = Td.CreateIndex()
        Idx.Name = "Pzn"
        Idx.Primary = False
        Idx.Unique = False
        Set IxFld = Idx.CreateField("Pzn")
        Idx.Fields.Append IxFld
        Td.Indexes.Append Idx
    
        Set Idx = Td.CreateIndex()
        Idx.Name = "Txt"
        Idx.Primary = False
        Idx.Unique = False
        Set IxFld = Idx.CreateField("Txt")
        Idx.Fields.Append IxFld
        Set IxFld = Idx.CreateField("Pzn")
        Idx.Fields.Append IxFld
        Set IxFld = Idx.CreateField("BestellDatum")
        Idx.Fields.Append IxFld
        Td.Indexes.Append Idx
    
        MerkzettelDB.TableDefs.Append Td
    
    
    'Tabelle Parameter
        Set Td = MerkzettelDB.CreateTableDef("Parameter")
    
        Set Fld = Td.CreateField("Name", dbText)
        Fld.AllowZeroLength = False
        Fld.Size = 30
        Td.Fields.Append Fld
    
        Set Fld = Td.CreateField("Wert", dbInteger)
        Td.Fields.Append Fld
    
        ' Indizes für Parameter
        Set Idx = Td.CreateIndex()
        Idx.Name = "Name"
        Idx.Primary = True
        Idx.Unique = False
        Set IxFld = Idx.CreateField("Name")
        Idx.Fields.Append IxFld
        Td.Indexes.Append Idx
    
        MerkzettelDB.TableDefs.Append Td
    
        MerkzettelDB.Close
    End If
    On Error GoTo DefErr
    
    OpenErg% = Merkzettel.OpenDatenbank("", False, False, MerkzettelDB)
    Set MerkzettelRec = MerkzettelDB.OpenRecordset("Merkzettel", dbOpenTable)
    Set MerkzettelParaRec = MerkzettelDB.OpenRecordset("Parameter", dbOpenTable)
    MerkzettelParaRec.Index = "Name"
    
    If (DbNeu%) Then
        BESORGT.OpenDatei
        
        DateiMax& = (BESORGT.DateiLen / BESORGT.RecordLen) - 1
        For i& = 1 To DateiMax&
            BESORGT.GetRecord (i& + 1)
            If (InStr("* ", BESORGT.flag) > 0) And (BESORGT.dat > 0) Then
                MerkzettelRec.AddNew
                
                MerkzettelRec!pzn = BESORGT.pzn
                MerkzettelRec!txt = BESORGT.text
                MerkzettelRec!BestellDatum = MakeDateStr$(BESORGT.dat)
                
                If (BESORGT.dt > 0) Then
                    MerkzettelRec!LieferDatum = MakeDateStr$(BESORGT.dt)
                Else
                    MerkzettelRec!LieferDatum = "01.01.1980"
                End If
                
                MerkzettelRec!Lief = BESORGT.Lief
                MerkzettelRec!Lm = BESORGT.Lm
                
                MerkzettelRec!loesch = 0
                
                MerkzettelRec.Update
            End If
        Next i&
        
        BESORGT.CloseDatei
    End If
Else
    On Error Resume Next
    Kill "merkzett.mdb"
    On Error GoTo DefErr
    BESORGT.OpenDatei
End If

OpenCreateMerkzettelDB% = ret%

Call DefErrPop
End Function

Function MakeDateStr$(iDat%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeDateStr$")
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
Dim ret$, h$

h$ = sDate(iDat%)

ret$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "."
If (Val(Mid$(h$, 5, 2)) > 50) Then
    ret$ = ret$ + "19"
Else
    ret$ = ret$ + "20"
End If
ret$ = ret$ + Mid$(h$, 5, 2)

MakeDateStr$ = ret$

Call DefErrPop
End Function
        


