Attribute VB_Name = "modSchwellwerte"
Option Explicit

Private Const DefErrModul = "SCHWELLWERTE.BAS"

Type BeobachtUmsaetzeType
    LiefAndDatum As String * 9
    WoTag As Byte
    Umsatz As Double
    UmsatzRabatt As Double
End Type
Public BeobachtUmsaetze() As BeobachtUmsaetzeType
Public AnzBeobachtUmsaetze%



Type SchwellLiefType
    lief As Integer
    rZeit As Integer
    rZeit60 As Integer
    SendungenAlt As Integer
    SendungenPlan As Integer
    AliquotProzent As Double
    MindAliquotProzent As Double
    AliquotMindestUmsatz As Double
    SendungenProTag(7) As Byte
    OrgBeobachtSendungen As Integer
    OrgBeobachtUmsatz(1) As Double
    BeobachtSendungen As Integer
    BeobachtUmsatz(1) As Double
    UmsatzProSendung(1) As Double
    UmsatzBisher(1) As Double
    UmsatzMitZuordnungen(1) As Double
    UmsatzMitMindest(1) As Double
    UmsatzMitSprung(1) As Double
    UmsatzMitBestLief(1) As Double
    PrognoseUmsatz(1) As Double
    PrognoseRabatt As Double
    PrognoseSchwellwert As Double
    MindestUmsatz As Double
    TabTyp As Byte
    RabattSprung As Integer
End Type
Public SchwellLief() As SchwellLiefType
Public AnzSchwellLief%
Public AktSchwellLief$



Type SchwellArtikelType
    zWert As Double
    DateiInd As Integer
    FlxInd As Integer
End Type
Public SchwellArtikel() As SchwellArtikelType
Public AnzSchwellArtikel%

Public SchwellwertAktiv%
Public SchwellwertMinuten%
Public SchwellwertSicherheit%
Public SchwellwertToleranz%
Public SchwellwertWarnungProz%
Public SchwellwertWarnungAb%
Public SchwellwertVorab%
Public SchwellwertGlaetten%

Public SchwellInfoSuch$
Public SchwellInfoName$

Public SchwellProt%

Public OptBelegung%()

Public OptimalAus%


Sub AuslesenSchwellwertArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenSchwellwertArtikel")
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
Dim i%, lief%, Max%, WasGeändert%, erg%, menge%, lfa%
Dim h$, h3$, Lac$, sRet$
Dim wwOrg$, wwNeu$

Call EinlesenWumsatzOk

frmAction.flxSortierung.Rows = 0

ww.SatzLock (1)
ww.GetRecord (1)

WasGeändert% = False
Max% = ww.erstmax
BekartCounter% = ww.erstcounter

AktUhrzeit% = Val(Format(Now, "HHMM"))
NaechstLiefernderLieferant% = HoleNaechstLiefernden%


Call frmAction.ZeigeSchwellwertAction("Auslesen der Bestell-Datei ...")

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    
    If (ww.status = 1) Then
    
        wwOrg$ = ww.RawData
        
        If (ww.OrgZeit = 0) Or (ww.IstSchwellArtikel) Then
            If (ww.OrgZeit = 0) Then ww.OrgZeit = SchwellLief(0).rZeit60
            If (ww.IstSchwellArtikel) Then
                ww.IstSchwellArtikel = 0
                ww.lief = 0
            End If
        End If

        Call EinzelSatz(1, Lac$)
    
        menge% = Abs(ww.bm)
        lfa% = ww.lief
    
        If ((ww.loesch = 0) And (ww.aktivlief = 0) And (ww.zukontrollieren <> "1")) Then
            sRet$ = CheckSchwellwert(i%, lfa%, Lac$, menge%)
            If (sRet$ <> "") Then
                h3$ = Format(ww.aep * menge%, "000000.00") + vbTab + Str$(i%) + vbTab + sRet$
                frmAction.flxSortierung.AddItem h3$
            End If
        End If
        
        wwNeu$ = ww.RawData
        If (wwOrg$ <> wwNeu$) And (ww.aktivlief = 0) Then
            ww.PutRecord (i% + 1)
            WasGeändert% = True
        End If
    End If
Next i%

Call SchwellwertArtikelZuordnen

Call SpeicherSchwellwertProtokoll

If (WasGeändert%) Then
    BekartCounter% = (BekartCounter% + 1) Mod 100
    ww.erstcounter = BekartCounter%
    ww.PutRecord (1)
End If

Call ww.SatzUnLock(1)

Call DefErrPop
End Sub

Function CheckSchwellwert$(dInd%, lfa%, Lac$, menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckSchwellwert$")
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
Dim i%, j%, k%, KontrollFlag%, iLief%, OkLief%(10), AnzOkLief%
Dim EK#, VK#, zWert#
Dim h$

CheckSchwellwert$ = ""

EK# = ww.aep
VK# = ww.AVP
zWert# = EK# * menge%
    
If (lfa% <> 0) Then
    Call AddUmsatz(lfa%, ww.pzn, ww.nnart, menge%, zWert#)
    Call DefErrPop: Exit Function
End If

If (ww.absage > 0) Then
    Call DefErrPop: Exit Function
End If

For j% = 0 To AnzZuordnungen% - 1
    KontrollFlag% = CheckBedingung%(Zuordnungen(j%).bedingung, EK#, VK#, menge%)
    
    If (KontrollFlag% = True) Then
        If (Zuordnungen(j%).lief(0) = 255) Then
            Call AddUmsatz(NaechstLiefernderLieferant%, ww.pzn, ww.nnart, menge%, zWert#)
            Call DefErrPop: Exit Function
        Else
            AnzOkLief% = 0
            For i% = 0 To 19
                iLief% = Zuordnungen(j%).lief(i%)
                If (iLief% > 0) Then
                    If (IstSchwellLieferant(iLief%) >= 0) Then
                        OkLief%(AnzOkLief%) = iLief%
                        AnzOkLief% = AnzOkLief% + 1
                    End If
                Else
                    Exit For
                End If
            Next i%
            If (AnzOkLief% = 1) Then
                Call AddUmsatz(OkLief%(0), ww.pzn, ww.nnart, menge%, zWert#)
            ElseIf (AnzOkLief% > 1) Then
'                h$ = ","
'                For k% = 1 To AnzOkLief%
'                    h$ = h$ + Mid$(Str$(OkLief%(k% - 1)), 2) + ","
'                Next k%
                h$ = String$(AnzSchwellLief%, "0")
                For k% = 1 To AnzOkLief%
'                    h$ = h$ + Mid$(Str$(OkLief%(k% - 1)), 2) + ","
                    iLief% = IstSchwellLieferant(OkLief%(k% - 1))
                    Mid$(h$, iLief% + 1, 1) = "1"
                Next k%
                CheckSchwellwert$ = PruefeZeitfenster$(dInd%, h$, ww.pzn, ww.nnart, menge%, zWert#)
            End If
        End If
        Call DefErrPop: Exit Function
    End If
Next j%

'h$ = AktSchwellLief$
h$ = String$(AnzSchwellLief%, "1")
CheckSchwellwert$ = PruefeZeitfenster$(dInd%, h$, ww.pzn, ww.nnart, menge%, zWert#)

Call DefErrPop
End Function

Sub AddUmsatz(lief%, pzn$, nnart, menge%, zWert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AddUmsatz")
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
Dim iLief%, iRabatt%, iNNart%
    
iNNart% = nnart
iLief% = IstSchwellLieferant(lief%)
If (iLief% >= 0) Then
    With SchwellLief(iLief%)
        .UmsatzMitZuordnungen(0) = .UmsatzMitZuordnungen(0) + zWert#
        iRabatt% = rabtab.HatRabatt(lief%, pzn$, iNNart%, Abs(menge%))
        If (iRabatt%) Then .UmsatzMitZuordnungen(1) = .UmsatzMitZuordnungen(1) + zWert#
        Print #SchwellProt%, "Z "; Format(lief%, "000"); Str$(Abs(iRabatt%)); " "; DruckZeile$
    End With
End If
                   
Call DefErrPop
End Sub

Sub EinlesenLieferantenAnrufe(Optional AutoInd% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenLieferantenAnrufe")
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
Dim i%, j%, k%, iLief%, rWoTag%, IstTag%, IstWoTag%
Dim h$, h2$, RestDat$, LetztSend$

Call frmAction.ZeigeSchwellwertAction("Anzahl erledigter/geplanter Sendungen für betroffene Lieferanten einlesen ...")

For j% = 0 To (AnzRufzeiten% - 1)
    iLief% = Rufzeiten(j%).Lieferant
    iLief% = lifzus.GetWumsatzLief(iLief%)
    iLief% = IstSchwellLieferant%(iLief%)
    If (iLief% >= 0) Then
        With SchwellLief(iLief%)
            For k% = 0 To 6
                rWoTag% = Rufzeiten(j%).WoTag(k%)
                If (rWoTag% > 0) Then
                    .SendungenProTag(rWoTag%) = .SendungenProTag(rWoTag%) + 1
                Else
                    Exit For
                End If
            Next k%
        End With
    End If
Next j%


IstTag% = Day(Now)
RestDat$ = Format(Now, "MMYY")
h$ = "01." + Format(Now, "MM.YYYY")
IstWoTag% = WeekDay(h$, vbMonday)

For i% = 1 To 31
    h$ = Format(i%, "00") + RestDat$
    If (iDate(h$) > 0) Then
        h2$ = Format(i%, "00") + "." + Format(Now, "MM.YYYY")
        If (IstFeiertag%(h2$) = False) Then
            For j% = 0 To (AnzRufzeiten% - 1)
                iLief% = Rufzeiten(j%).Lieferant
                iLief% = lifzus.GetWumsatzLief(iLief%)
                iLief% = IstSchwellLieferant%(iLief%)
                If (iLief% >= 0) Then
                    With SchwellLief(iLief%)
                        For k% = 0 To 6
                            rWoTag% = Rufzeiten(j%).WoTag(k%)
                            If (rWoTag% = IstWoTag%) Then
                                If (i% < IstTag%) Then
                                    .SendungenAlt = .SendungenAlt + 1
                                ElseIf (i% = IstTag%) Then
                                    LetztSend$ = Format(Rufzeiten(j%).LetztSend, "00000000")
                                    LetztSend$ = Left$(LetztSend$, 4) + Right$(LetztSend$, 2)
                                    If (LetztSend$ = h$) And (j% <> AutoInd%) Then
                                        .SendungenAlt = .SendungenAlt + 1
                                    Else
                                        .SendungenPlan = .SendungenPlan + 1
                                    End If
                                Else
                                    .SendungenPlan = .SendungenPlan + 1
                                End If
                                Exit For
                            ElseIf (rWoTag% <= 0) Then
                                Exit For
                            End If
                        Next k%
                    End With
                End If
            Next j%
        End If
    End If
    
    IstWoTag% = IstWoTag% + 1
    If (IstWoTag% > 7) Then
        IstWoTag% = 1
    End If
Next i%


Call DefErrPop
End Sub

Sub OptimaleAuswahl(SollWert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OptimaleAuswahl%")
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
Dim i%, ActIndex%, WertIndex%, MaxIndex%, ActBelegung%(), Weiterschieben%
Dim ActWert#, OptWert#

MaxIndex% = UBound(SchwellArtikel)
ReDim ActBelegung%(MaxIndex%)
ReDim OptBelegung%(MaxIndex%)

OptWert# = 0#
For i% = 0 To MaxIndex%
    OptBelegung%(i%) = i%
    OptWert# = OptWert# + SchwellArtikel(i%).zWert
Next i%

If (OptWert# > SollWert#) Then

    frmAction.tmrOptimal.Interval = 2000
    frmAction.tmrOptimal.Enabled = True
    OptimalAus% = False

    ActIndex% = 0
    ActBelegung%(ActIndex%) = 0
    ActWert# = SchwellArtikel(ActBelegung%(ActIndex%)).zWert
    Do
    '    ActWert# = 0
    '    For i% = 0 To ActIndex%
    '        ActWert# = ActWert# + werte#(ActBelegung%(i%))
    '    Next i%
            
    '    For i% = 0 To ActIndex%
    '        Debug.Print Str$(ActBelegung%(i%));
    '    Next i%
    '    Debug.Print Str$(OptWert#);
    '    Debug.Print
    
        If (OptimalAus%) Then Exit Do
        DoEvents
        
        Weiterschieben% = True
        If (ActWert# >= SollWert#) Then
            'eine Lösung gefunden
            If (ActWert# < OptWert#) Then
                'ist bisher beste Lösung
                OptWert# = ActWert#
                For i% = 0 To ActIndex%
                    OptBelegung%(i%) = ActBelegung%(i%)
                Next i%
                For i% = (ActIndex% + 1) To MaxIndex%
                    OptBelegung%(i%) = -1
                Next i%
            End If
        Else
            'Sollwert noch nicht erreicht
            If (ActBelegung%(ActIndex%) < MaxIndex%) Then
                'man kann hinten noch was dazuhängen
                ActIndex% = ActIndex% + 1
                ActBelegung%(ActIndex%) = ActBelegung%(ActIndex% - 1) + 1
                ActWert# = ActWert# + SchwellArtikel(ActBelegung%(ActIndex%)).zWert
                Weiterschieben% = False
            End If
        End If
                
        If (Weiterschieben%) Then
            'bisher vorletzten Pointer auf nächsten Wert schieben
            If (ActIndex% = 0) Then Exit Do
            ActWert# = ActWert# - SchwellArtikel(ActBelegung%(ActIndex%)).zWert
            ActIndex% = ActIndex% - 1
            ActWert# = ActWert# - SchwellArtikel(ActBelegung%(ActIndex%)).zWert
            ActBelegung%(ActIndex%) = ActBelegung%(ActIndex%) + 1
            ActWert# = ActWert# + SchwellArtikel(ActBelegung%(ActIndex%)).zWert
        End If
    Loop
    
    frmAction.tmrOptimal.Enabled = False
End If

For i% = 0 To MaxIndex%
    If (OptBelegung%(i%) = -1) Then
        ReDim Preserve OptBelegung%(i% - 1)
        Exit For
    End If
Next i%

Call DefErrPop
End Sub

Sub InitWumsatzDat()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitWumsatzDat")
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
Dim i%, WUMSATZ%
Dim buf$

If (Dir("WUMSATZ.DAT") = "") Then
    WUMSATZ% = FileOpen("WUMSATZ.DAT", "O")
    buf$ = Space$(10)
    For i% = 0 To 2
        Mid$(buf$, 2 + (i% * 3), 3) = Format(i% + 1, "000")
    Next i%
    Print #WUMSATZ%, buf$
    Close #WUMSATZ%
End If

Call DefErrPop
End Sub

Sub EinlesenWumsatzOk()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWumsatzOk")
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
Dim i%, j%, WUMSATZ%, tlief%, tmpDiff%, fehler%, AnzBeobachtTage%, gef%, ind%
Dim AktBeobachtSendungen%, RabattFlag%, anzsend%, AnzRowsWeg%
Dim tValue#, ums#, ums2#, MindAliquotProzent#, schnitt#, ToleranzUms#
Dim buf$, SollStr$, h$, h2$
Dim tdatum As Date, AbDatum As Date


AnzBeobachtUmsaetze% = 0
AnzBeobachtTage% = para.bp1 * para.iBP
If (AnzBeobachtTage% < 31) Then AnzBeobachtTage% = 31
AbDatum = Date - AnzBeobachtTage% - 1

If (Dir$("WUMSATZ.DAT") <> "") Then
    WUMSATZ% = FileOpen("WUMSATZ.DAT", "I")
    Do While (EOF(WUMSATZ%) = False)
        Line Input #WUMSATZ%, buf$
    
        fehler% = 0
        On Error GoTo ErrorWumsatz2
        tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
        On Error GoTo DefErr
        If (fehler% = 0) Then
            If (tdatum < AbDatum) Then fehler% = 1
        End If
        If (fehler% = 0) Then
            tlief% = Val(Mid$(buf$, 8, 3))
    '        If (tLief% <= 0) Or (tLief% > 255) Then fehler% = 1
            tlief% = IstSchwellLieferant%(tlief%)
            If (tlief% < 0) Then fehler% = 1
        End If
        If (fehler% = 0) Then
            tValue# = Val(Mid$(buf$, 12, 9)) / 100#
            RabattFlag% = (Mid$(buf$, 21, 1) <> "*")
            
            If (tdatum < Date) Then
                SollStr$ = Format(SchwellLief(tlief%).lief, "000") + Left$(buf$, 6)
                gef% = False
                For i% = 0 To (AnzBeobachtUmsaetze% - 1)
                    With BeobachtUmsaetze(i%)
                        If (.LiefAndDatum = SollStr$) Then
                            .Umsatz = .Umsatz + tValue#
                            If (RabattFlag%) Then .UmsatzRabatt = .UmsatzRabatt + tValue#
                            gef% = True
                            Exit For
                        End If
                    End With
                Next i%
                If (gef% = False) Then
                    i% = AnzBeobachtUmsaetze%
                    ReDim Preserve BeobachtUmsaetze(i%)
                    With BeobachtUmsaetze(i%)
                        .LiefAndDatum = SollStr$
                        .WoTag = WeekDay(tdatum, vbMonday)
                        .Umsatz = tValue#
                        If (RabattFlag%) Then
                            .UmsatzRabatt = tValue#
                        Else
                            .UmsatzRabatt = 0#
                        End If
                    End With
                    AnzBeobachtUmsaetze% = i% + 1
                End If
            End If
        
            tmpDiff% = DateDiff("M", tdatum, Date)
            If (tmpDiff% = 0) Then
                With SchwellLief(tlief%)
                    .UmsatzBisher(0) = .UmsatzBisher(0) + tValue#
                    If (RabattFlag%) Then .UmsatzBisher(1) = .UmsatzBisher(1) + tValue#
                End With
            End If
        End If
    Loop
    Close WUMSATZ%
End If


For i% = 0 To (AnzSchwellLief% - 1)
    With SchwellLief(i%)
        SollStr$ = Format(.lief, "000")
                    
        frmAction.flxSortierung.Rows = 0

        .OrgBeobachtSendungen = 0
        .OrgBeobachtUmsatz(0) = 0#
        .OrgBeobachtUmsatz(1) = 0#
        For j% = 0 To (AnzBeobachtUmsaetze% - 1)
            If (Left$(BeobachtUmsaetze(j%).LiefAndDatum, 3) = SollStr$) Then
                ums# = BeobachtUmsaetze(j%).Umsatz
                ums2# = BeobachtUmsaetze(j%).UmsatzRabatt
                anzsend% = .SendungenProTag(BeobachtUmsaetze(j%).WoTag)
                If (anzsend% = 0) Then anzsend% = 1
                schnitt# = ums# / anzsend%
                h$ = Format(ums#, "000000.00") + vbTab
                h$ = h$ + Format(ums2#, "000000.00") + vbTab
                h$ = h$ + Format(anzsend%, "0") + vbTab
                h$ = h$ + Format(schnitt#, "000000.00")
                frmAction.flxSortierung.AddItem h$
                
'               If (.lief = 1) Then Debug.Print BeobachtUmsaetze(j%).LiefAndDatum + " " + h$
            
                .OrgBeobachtSendungen = .OrgBeobachtSendungen + anzsend%
                .OrgBeobachtUmsatz(0) = .OrgBeobachtUmsatz(0) + ums#
                .OrgBeobachtUmsatz(1) = .OrgBeobachtUmsatz(1) + ums2#
            End If
        Next j%
        
        anzsend% = 0: ums# = 0#: ums2# = 0#
        If (SchwellwertGlaetten%) And (frmAction.flxSortierung.Rows > 0) Then
            With frmAction.flxSortierung
                .row = 0
                .col = 3
                .RowSel = .Rows - 1
                .ColSel = 3
                .Sort = 5
                .col = 0
                .ColSel = .Cols - 1
                
                AnzRowsWeg% = Int(.Rows / 10 + 0.5)
                
'                If (SchwellLief(i%).lief = 1) Then
'                    Debug.Print Str$(AnzRowsWeg%)
'                    For j% = 1 To .Rows
'                        Debug.Print .TextMatrix(j% - 1, 3)
'                    Next j%
'                End If
                
                For j% = 1 To AnzRowsWeg%
                    .RemoveItem 0
                Next j%
                For j% = 1 To AnzRowsWeg%
                    .RemoveItem .Rows - 1
                Next j%
            
                For j% = 0 To (.Rows - 1)
                    ums# = ums# + CDbl(.TextMatrix(j%, 0))
                    ums2# = ums2# + CDbl(.TextMatrix(j%, 1))
                    anzsend% = anzsend% + Val(.TextMatrix(j%, 2))
                Next j%
            End With
        End If
        If (anzsend% = 0) Then
            anzsend% = .OrgBeobachtSendungen
            ums# = .OrgBeobachtUmsatz(0)
            ums2# = .OrgBeobachtUmsatz(1)
        End If
        .BeobachtSendungen = anzsend%
        .BeobachtUmsatz(0) = ums#
        .BeobachtUmsatz(1) = ums2#
        
        
        If (.BeobachtSendungen > 0) Then
            .UmsatzProSendung(0) = .BeobachtUmsatz(0) / .BeobachtSendungen
            .UmsatzProSendung(1) = .BeobachtUmsatz(1) / .BeobachtSendungen
        Else
            .UmsatzProSendung(0) = 0#
            .UmsatzProSendung(1) = 0#
        End If
        
        .UmsatzMitZuordnungen(0) = .UmsatzBisher(0)
        .UmsatzMitZuordnungen(1) = .UmsatzBisher(1)
        
        lifzus.GetRecord (.lief + 1)
        .MindestUmsatz = lifzus.MindestUmsatz
                
        If (.SendungenAlt + .SendungenPlan > 0) Then
            .AliquotProzent = CDbl(.SendungenAlt) / (.SendungenAlt + .SendungenPlan)
            .MindAliquotProzent = CDbl(.SendungenAlt + 1) / (.SendungenAlt + .SendungenPlan)
        Else
            .AliquotProzent = 0#
            .MindAliquotProzent = 0#
        End If
        .AliquotMindestUmsatz = .MindestUmsatz * .MindAliquotProzent
    
        .PrognoseUmsatz(0) = .UmsatzBisher(0) + (.SendungenPlan * .UmsatzProSendung(0))
        .PrognoseUmsatz(1) = .UmsatzBisher(1) + (.SendungenPlan * .UmsatzProSendung(1))
        
        ToleranzUms# = .PrognoseUmsatz(1) / 100# * (100# + SchwellwertToleranz)
        .PrognoseSchwellwert = lifzus.GetPrognoseSchwellwert(ToleranzUms#)
        .PrognoseRabatt = lifzus.GetPrognoseRabatt(ToleranzUms#)
        .TabTyp = lifzus.TabTyp
    
    End With
Next i%

Call DefErrPop: Exit Sub
    
ErrorWumsatz2:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

Function IstSchwellLieferant%(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstSchwellLieferant%")
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

ret% = -1
For i% = 0 To (AnzSchwellLief% - 1)
    If (SchwellLief(i%).lief = iLief%) Then
        ret% = i%
        Exit For
    End If
Next i%

IstSchwellLieferant% = ret%

Call DefErrPop
End Function

Function PruefeZeitfenster$(dInd%, sLiefs$, pzn$, nnart, menge%, zWert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeZeitfenster$")
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
Dim i%, ind%, iLief%, tlief%, anz%, retLief%, iNNart%
Dim h$, ret$

ret$ = ""
anz% = 0
retLief% = 0
iNNart% = nnart

'h$ = Mid$(sLiefs$, 2)
'Do
'    ind% = InStr(h$, ",")
'    If (ind% <= 0) Then Exit Do
'    iLief% = Val(Left$(h$, ind% - 1))
'
'    tlief% = IstSchwellLieferant(iLief%)
'    If (tlief% >= 0) Then
'        If (SchwellwertMinuten% = 0) Or ((ww.OrgZeit + SchwellwertMinuten%) >= SchwellLief(tlief%).rZeit60) Then
'            If (ret$ = "") Then ret$ = ","
'            ret$ = ret$ + Mid$(Str$(iLief%), 2) + ","
'
'            anz% = anz% + 1
'            retLief% = iLief%
'        End If
'    End If
'
'    h$ = Mid$(h$, ind% + 1)
'Loop

For i% = 0 To (AnzSchwellLief% - 1)
    If (Mid$(sLiefs$, i% + 1, 1) = "1") Then
        If (SchwellwertMinuten% = 0) Or ((ww.OrgZeit + SchwellwertMinuten%) >= SchwellLief(i%).rZeit60) Then
            If (ret$ = "") Then ret$ = String$(AnzSchwellLief%, "0")
            Mid$(ret$, i% + 1, 1) = "1"
            anz% = anz% + 1
            retLief% = SchwellLief(i%).lief
        End If
    End If
Next i%


If (anz% = 0) Then
    For i% = 0 To (AnzSchwellLief% - 1)
'        h$ = "," + Mid$(Str$(SchwellLief(i%).lief), 2) + ","
'        ind% = InStr(sLiefs$, h$)
'        If (ind% > 0) Then
        If (Mid$(sLiefs$, i% + 1, 1) = "1") Then
            anz% = 1
            retLief% = SchwellLief(i%).lief
            Exit For
        End If
    Next i%
End If

If (anz% = 1) Then
    ret$ = ""
    ww.IstSchwellArtikel = 3
    ww.lief = retLief%
    ww.PutRecord (dInd% + 1)
    Call AddUmsatz(retLief%, pzn$, iNNart%, menge%, zWert#)
End If

PruefeZeitfenster$ = ret$

Call DefErrPop
End Function

Sub SpeicherSchwellwertDaten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherSchwellwertDaten")
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
Dim i%, j%

Call HoleIniRufzeiten
Call EinlesenLieferantenAnrufe
Call EinlesenWumsatzOk

For j% = 0 To (AnzSchwellLief% - 1)
    With SchwellLief(j%)
        lifzus.GetRecord (.lief + 1)
        
        lifzus.OrgBeobachtSendungen = .OrgBeobachtSendungen
        For i% = 0 To 1
            lifzus.OrgBeobachtUmsatz(i%) = .OrgBeobachtUmsatz(i%)
        Next i%
        lifzus.BeobachtSendungen = .BeobachtSendungen
        For i% = 0 To 1
            lifzus.BeobachtUmsatz(i%) = .BeobachtUmsatz(i%)
        Next i%
        For i% = 0 To 1
            lifzus.UmsatzProSendung(i%) = .UmsatzProSendung(i%)
        Next i%
        
        lifzus.SendungenAlt = .SendungenAlt
        lifzus.SendungenPlan = .SendungenPlan
        lifzus.AliquotProzent = .AliquotProzent
        
        For i% = 0 To 1
            lifzus.UmsatzBisher(i%) = .UmsatzBisher(i%)
        Next i%
        For i% = 0 To 1
            lifzus.PrognoseUmsatz(i%) = .PrognoseUmsatz(i%)
        Next i%
        lifzus.PrognoseSchwellwert = .PrognoseSchwellwert
        lifzus.PrognoseRabatt = .PrognoseRabatt ' lifzus.GetPrognoseRabatt(.PrognoseUmsatz(1))
        
        lifzus.PutRecord (.lief + 1)
    End With
Next j%

Call DefErrPop
End Sub
                  
Sub SchwellwertArtikelZuordnen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SchwellwertArtikelZuordnen")
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
Dim i%, j%, k%, iRabatt%, GibtsNochWelche%, DatInd%
Dim BestRabatt#, DiffUmsatz#, ToleranzUms#
Dim PrognoseUmsatzOhne#, PrognoseUmsatzMit#, UmsatzDazu#
Dim PrognoseSchwellwert#, AliquotSchwellwert#, PrognoseRabattOhne#, PrognoseRabattMit#
Dim h3$

For k% = 0 To (AnzSchwellLief% - 1)
    With SchwellLief(k%)
        For i% = 0 To 1
            .UmsatzMitMindest(i%) = .UmsatzMitZuordnungen(i%)
            .UmsatzMitSprung(i%) = .UmsatzMitMindest(i%)
            .UmsatzMitBestLief(i%) = .UmsatzMitSprung(i%)
        Next i%
    End With
Next k%

If (frmAction.flxSortierung.Rows > 0) Then
    With frmAction.flxSortierung
        .row = 0
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = 0
        .Sort = 5
        .col = 0
        .ColSel = .Cols - 1
    End With

    ReDim SchwellArtikel(0)
        
    For k% = 0 To (AnzSchwellLief% - 1)
        
        With SchwellLief(k%)
            
            If (.MindestUmsatz > 0#) Then
    
                DiffUmsatz# = .AliquotMindestUmsatz - .UmsatzMitZuordnungen(0)
                
                If (DiffUmsatz# > 0#) Then
                    AnzSchwellArtikel% = 0
                    For i% = 0 To (frmAction.flxSortierung.Rows - 1)
                        h3$ = frmAction.flxSortierung.TextMatrix(i%, 2)
                        If (Mid$(h3$, k% + 1, 1) = "1") Then
                            j% = AnzSchwellArtikel%
                            ReDim Preserve SchwellArtikel(j%)
                            SchwellArtikel(j%).zWert = CDbl(frmAction.flxSortierung.TextMatrix(i%, 0))
                            SchwellArtikel(j%).DateiInd = Val(frmAction.flxSortierung.TextMatrix(i%, 1))
                            SchwellArtikel(j%).FlxInd = i%
                            AnzSchwellArtikel% = j% + 1
                        End If
                    Next i%
                    
                    If (AnzSchwellArtikel% > 0) Then
                        Call frmAction.ZeigeSchwellwertAction("Lieferant" + Str$(.lief) + " auf aliquoten Mindestumsatz auffüllen ...")

                        Call OptimaleAuswahl(DiffUmsatz#)
                        For i% = 0 To UBound(OptBelegung%)
                            j% = OptBelegung%(i%)
                            .UmsatzMitMindest(0) = .UmsatzMitMindest(0) + SchwellArtikel(j%).zWert
                            
                            ww.GetRecord (SchwellArtikel(j%).DateiInd + 1)
        
                            iRabatt% = rabtab.HatRabatt(.lief, ww.pzn, ww.nnart, Abs(ww.bm))
                            If (iRabatt%) Then
                                .UmsatzMitMindest(1) = .UmsatzMitMindest(1) + SchwellArtikel(j%).zWert
                            End If
                            
                            Print #SchwellProt%, "M "; Format(.lief, "000"); Str$(Abs(iRabatt%)); " "; DruckZeile$
                            ww.lief = .lief
                            ww.IstSchwellArtikel = 1
                            ww.PutRecord (SchwellArtikel(j%).DateiInd + 1)
                            
                            frmAction.flxSortierung.TextMatrix(SchwellArtikel(j%).FlxInd, 2) = "XXXXXXX"
'                            WasGeändert% = True
                        Next i%
                    End If
                    
                End If
                
            End If
            
            .UmsatzMitSprung(0) = .UmsatzMitMindest(0)
            .UmsatzMitSprung(1) = .UmsatzMitMindest(1)
            .UmsatzMitBestLief(0) = .UmsatzMitMindest(0)
            .UmsatzMitBestLief(1) = .UmsatzMitMindest(1)
        End With
    Next k%
    
    GibtsNochWelche% = False
    With frmAction.flxSortierung
        For i% = 0 To (.Rows - 1)
            If (.TextMatrix(i%, 2) <> "XXXXXXX") Then
                GibtsNochWelche% = True
                Exit For
            End If
        Next i%
    End With
    
    If (GibtsNochWelche%) Then
                    
        Call frmAction.ZeigeSchwellwertAction("Ermitteln möglicher Schwellwert-Sprünge ...")
        
        For k% = 0 To (AnzSchwellLief% - 1)
            
            With SchwellLief(k%)
        
                lifzus.GetRecord (.lief + 1)
                
                .RabattSprung = False
                
                PrognoseUmsatzOhne# = .UmsatzMitMindest(1) + ((.SendungenPlan - 1) * .UmsatzProSendung(1))
                UmsatzDazu# = 0#
        
                ToleranzUms# = PrognoseUmsatzOhne# / 100# * (100# + SchwellwertToleranz)
                PrognoseRabattOhne# = lifzus.GetPrognoseRabatt(ToleranzUms#)
           
                For i% = 0 To (frmAction.flxSortierung.Rows - 1)
                    h3$ = frmAction.flxSortierung.TextMatrix(i%, 2)
                    If (Mid$(h3$, k% + 1, 1) = "1") Then
                        UmsatzDazu# = UmsatzDazu# + CDbl(frmAction.flxSortierung.TextMatrix(i%, 0)) ' SchwellArtikel(j%).zWert
                    End If
                Next i%
                
                If (UmsatzDazu# > 0#) Then
                    PrognoseUmsatzMit# = PrognoseUmsatzOhne# + UmsatzDazu#
                    ToleranzUms# = PrognoseUmsatzMit# / 100# * (100# + SchwellwertToleranz)
                    PrognoseRabattMit# = lifzus.GetPrognoseRabatt(ToleranzUms#)
                    If (PrognoseRabattMit# > PrognoseRabattOhne#) Then
                        .RabattSprung = True
                    End If
                End If
                        
            End With
        Next k%
        
        

        For k% = 0 To (AnzSchwellLief% - 1)
            
            With SchwellLief(k%)
        
                If (.RabattSprung) Then
                
                    lifzus.GetRecord (.lief + 1)
                
                    PrognoseUmsatzOhne# = .UmsatzMitMindest(1) + ((.SendungenPlan - 1) * .UmsatzProSendung(1))
                    UmsatzDazu# = 0#
            
                    ToleranzUms# = PrognoseUmsatzOhne# / 100# * (100# + SchwellwertToleranz)
                    PrognoseRabattOhne# = lifzus.GetPrognoseRabatt(ToleranzUms#)
               
                    AnzSchwellArtikel% = 0
                    For i% = 0 To (frmAction.flxSortierung.Rows - 1)
                        h3$ = frmAction.flxSortierung.TextMatrix(i%, 2)
                        If (Mid$(h3$, k% + 1, 1) = "1") Then
                            DatInd% = Val(frmAction.flxSortierung.TextMatrix(i%, 1))
                            ww.GetRecord (DatInd% + 1)
        
                            iRabatt% = rabtab.HatRabatt(.lief, ww.pzn, ww.nnart, Abs(ww.bm))
                            If (iRabatt%) Then
                                j% = AnzSchwellArtikel%
                                ReDim Preserve SchwellArtikel(j%)
                                SchwellArtikel(j%).zWert = CDbl(frmAction.flxSortierung.TextMatrix(i%, 0))
                                SchwellArtikel(j%).DateiInd = Val(frmAction.flxSortierung.TextMatrix(i%, 1))
                                SchwellArtikel(j%).FlxInd = i%
                                AnzSchwellArtikel% = j% + 1
                                UmsatzDazu# = UmsatzDazu# + SchwellArtikel(j%).zWert
                            End If
                        End If
                    Next i%
                    
                    If (UmsatzDazu# > 0#) Then
                        PrognoseUmsatzMit# = PrognoseUmsatzOhne# + UmsatzDazu#
                        ToleranzUms# = PrognoseUmsatzMit# / 100# * (100# + SchwellwertToleranz)
                        PrognoseRabattMit# = lifzus.GetPrognoseRabatt(ToleranzUms#)
                        If (PrognoseRabattMit# > PrognoseRabattOhne#) Then
                            PrognoseSchwellwert# = lifzus.GetPrognoseSchwellwert(ToleranzUms#)
                            AliquotSchwellwert# = PrognoseSchwellwert# * .MindAliquotProzent
                            DiffUmsatz# = AliquotSchwellwert# - .UmsatzMitMindest(1)
                
                            Call frmAction.ZeigeSchwellwertAction("Lieferant" + Str$(.lief) + " auf aliquoten Schwellwert auffüllen ...")
                            Call OptimaleAuswahl(DiffUmsatz#)
                            For i% = 0 To UBound(OptBelegung%)
                                j% = OptBelegung%(i%)
                                .UmsatzMitSprung(0) = .UmsatzMitBestLief(0) + SchwellArtikel(j%).zWert
                                .UmsatzMitSprung(1) = .UmsatzMitBestLief(1) + SchwellArtikel(j%).zWert
                                iRabatt% = True
                                
                                ww.GetRecord (SchwellArtikel(j%).DateiInd + 1)
                                
                                Print #SchwellProt%, "P "; Format(.lief, "000"); Str$(Abs(iRabatt%)); " "; DruckZeile$
                                ww.lief = .lief
                                ww.IstSchwellArtikel = 4
                                ww.PutRecord (SchwellArtikel(j%).DateiInd + 1)
                                
                                frmAction.flxSortierung.TextMatrix(SchwellArtikel(j%).FlxInd, 2) = "XXXXXXX"
                            Next i%
                        End If
                    End If
                    
                    .UmsatzMitBestLief(0) = .UmsatzMitSprung(0)
                    .UmsatzMitBestLief(1) = .UmsatzMitSprung(1)
        
                End If
                        
            End With
        Next k%
    
    End If
        
    GibtsNochWelche% = False
    With frmAction.flxSortierung
        For i% = 0 To (.Rows - 1)
            If (.TextMatrix(i%, 2) <> "XXXXXXX") Then
                GibtsNochWelche% = True
                Exit For
            End If
        Next i%
    End With
    
    If (GibtsNochWelche%) Then
    
        For k% = 0 To (AnzSchwellLief% - 1)
            With SchwellLief(k%)
                Call rabtab.InitBestLief(k%, .lief)
            End With
        Next k%
            
        ReDim SchwellArtikel(0)
        
        With frmAction.flxSortierung
            Call frmAction.ZeigeSchwellwertAction("Ermitteln des günstigsten Lieferanten für restliche Artikel ...")
            
            For i% = 0 To (.Rows - 1)
                h3$ = .TextMatrix(i%, 2)
                If (h3$ <> "XXXXXXX") Then
                    SchwellArtikel(0).zWert = CDbl(.TextMatrix(i%, 0))
                    SchwellArtikel(0).DateiInd = Val(.TextMatrix(i%, 1))
                    SchwellArtikel(0).FlxInd = i%
                    
                    
                    ww.GetRecord (SchwellArtikel(0).DateiInd + 1)
                    
                    k% = rabtab.GetBestLief(h3$, ww.pzn, ww.nnart, Abs(ww.bm), BestRabatt#)
                                        
                    SchwellLief(k%).UmsatzMitBestLief(0) = SchwellLief(k%).UmsatzMitBestLief(0) + SchwellArtikel(0).zWert
                    
                    iRabatt% = rabtab.HatRabatt(SchwellLief(k%).lief, ww.pzn, ww.nnart, Abs(ww.bm))
                    If (iRabatt%) Then
                        SchwellLief(k%).UmsatzMitBestLief(1) = SchwellLief(k%).UmsatzMitBestLief(1) + SchwellArtikel(0).zWert
                    End If
                            
                    Print #SchwellProt%, "B "; Format(SchwellLief(k%).lief, "000"); Str$(Abs(iRabatt%)); " "; DruckZeile$
                    ww.lief = SchwellLief(k%).lief
                    ww.IstSchwellArtikel = 2
                    ww.PutRecord (SchwellArtikel(0).DateiInd + 1)
                    
'                    WasGeändert% = True
                End If
            Next i%
        End With
        
    End If
End If

Call DefErrPop
End Sub

Sub SpeicherSchwellwertProtokoll()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherSchwellwertProtokoll")
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
Dim i%, k%
Dim diff#
Dim h$, h3$

Call frmAction.ZeigeSchwellwertAction("Speichern des Schwellwert-Protokolls ...")

Print #SchwellProt%, "A"; Format(AnzSchwellLief%, "0")
For k% = 0 To (AnzSchwellLief% - 1)
    With SchwellLief(k%)
        Print #SchwellProt%, "*"; Format(.lief, "0")
        
        h3$ = "A " + Format(.lief, "000 ")
        h$ = Format(.rZeit, "0"): Print #SchwellProt%, h3$ + Left$(h$, 2) + ":" + Mid$(h$, 3)
        Print #SchwellProt%, h3$ + " "
        
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzProSendung(i%), "# ### ##0.00")
        Next i%
        
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.PrognoseUmsatz(i%), "# ### ##0.00")
        Next i%
        
        'Prognose Rabatt
'        Print #SchwellProt%, Format(.PrognoseRabatt, "0.00")
        
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzBisher(i%), "# ### ##0.00")
        Next i%
        For i% = 0 To 1
            diff# = .UmsatzMitZuordnungen(i%) - .UmsatzBisher(i%)
            If (diff# <> 0#) Then
                Print #SchwellProt%, h3$ + Format(diff#, "# ### ##0.00")
            Else
                Print #SchwellProt%, h3$ + " "
            End If
        Next i%
        For i% = 0 To 1
            diff# = .UmsatzMitMindest(i%) - .UmsatzMitZuordnungen(i%)
            If (diff# <> 0#) Then
                Print #SchwellProt%, h3$ + Format(diff#, "# ### ##0.00")
            Else
                Print #SchwellProt%, h3$ + " "
            End If
        Next i%
        For i% = 0 To 1
            diff# = .UmsatzMitSprung(i%) - .UmsatzMitMindest(i%)
            If (diff# <> 0#) Then
                Print #SchwellProt%, h3$ + Format(diff#, "# ### ##0.00")
            Else
                Print #SchwellProt%, h3$ + " "
            End If
        Next i%
        For i% = 0 To 1
            diff# = .UmsatzMitBestLief(i%) - .UmsatzMitSprung(i%)
            If (diff# <> 0#) Then
                Print #SchwellProt%, h3$ + Format(diff#, "# ### ##0.00")
            Else
                Print #SchwellProt%, h3$ + " "
            End If
        Next i%
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzMitBestLief(i%), "# ### ##0.00")
        Next i%
        
        h3$ = "Q " + Format(.lief, "000 ")
        diff# = .AliquotMindestUmsatz - .UmsatzMitZuordnungen(0)
        Print #SchwellProt%, h3$ + Format(diff#, "# ### ##0.00")
        
        h3$ = "S " + Format(.lief, "000 ")
        Print #SchwellProt%, h3$ + Format(.OrgBeobachtSendungen, "0")
        Print #SchwellProt%, h3$ + " "
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.OrgBeobachtUmsatz(i%), "# ### ##0.00")
        Next i%
        Print #SchwellProt%, h3$ + Format(.BeobachtSendungen, "0")
        Print #SchwellProt%, h3$ + " "
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.BeobachtUmsatz(i%), "# ### ##0.00")
        Next i%
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzProSendung(i%), "# ### ##0.00")
        Next i%

        
        h3$ = "U " + Format(.lief, "000 ")
        Print #SchwellProt%, h3$ + Format(.SendungenAlt, "0")
        Print #SchwellProt%, h3$ + " "
        Print #SchwellProt%, h3$ + Format(.SendungenPlan, "0")
        Print #SchwellProt%, h3$ + " "
        Print #SchwellProt%, h3$ + Format(Int(.AliquotProzent * 100# + 0.5), "0")
        Print #SchwellProt%, h3$ + " "
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzProSendung(i%), "# ### ##0.00")
        Next i%
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.UmsatzBisher(i%), "# ### ##0.00")
        Next i%
        For i% = 0 To 1
            Print #SchwellProt%, h3$ + Format(.PrognoseUmsatz(i%), "# ### ##0.00")
        Next i%
        
        If (k% = 0) And (.SendungenPlan <= SchwellwertWarnungAb%) Then Call CheckSchwellwertWarnung
        
    End With
Next k%

Call DefErrPop
End Sub

Sub CheckSchwellwertWarnung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckSchwellwertWarnung")
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
Dim i%, iWarnung%, iProz%, WARN_HANDLE%
Dim SchwellUms#
Dim h$, proz$(1)
            
iWarnung% = False
proz$(0) = "   "
proz$(1) = "   "

With SchwellLief(0)
    If (.MindestUmsatz > 0#) Then
    
        If (.SendungenPlan > 0) Then .SendungenPlan = .SendungenPlan - 1
        For i% = 0 To 1
            .PrognoseUmsatz(i%) = .UmsatzMitBestLief(i%) + (.SendungenPlan * .UmsatzProSendung(i%))
        Next i%
        
        iProz% = Int((.PrognoseUmsatz(0) / .MindestUmsatz) * 100# + 0.5)
        If (iProz% <= SchwellwertWarnungProz%) Then
            iWarnung% = True
            proz$(0) = Format(iProz%, "000")
        End If
    End If
        
    If (.TabTyp = 0) Then
        SchwellUms# = .PrognoseSchwellwert
        If (SchwellUms# = 0#) Then
            iProz% = 100
        Else
            iProz% = Int((.PrognoseUmsatz(1) / SchwellUms#) * 100# + 0.5)
        End If
        If (iProz% <= SchwellwertWarnungProz%) Then
            iWarnung% = True
            proz$(1) = Format(iProz%, "000")
        End If
    End If
    
    If (iWarnung%) Then
        lif.GetRecord (.lief + 1)
        h$ = Trim$(lif.kurz)
        If (h$ = String$(Len(h$), 0)) Then h$ = ""
        If (h$ = "") Then
            h$ = "(" + Str$(.lief) + ")"
        End If
        h$ = Format(Now, "dd.mm.yyyy hh:nn ") + h$
        
        For i% = 0 To 1
            h$ = h$ + " " + proz$(i%)
        Next i%
        
        WARN_HANDLE% = FreeFile
        Open "WINWARN.TXT" For Append As #WARN_HANDLE%
        Print #WARN_HANDLE%, h$
        Close #WARN_HANDLE%
    End If
End With

Call DefErrPop
End Sub

Sub ShowSchwellwertWarnung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowSchwellwertWarnung")
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
Dim i%, AnzWarn%, WARN_HANDLE%
Dim h$, h2$, h3$, WarnStr$()
            
AnzWarn% = 0
            
On Error Resume Next
WARN_HANDLE% = FreeFile
Open "WINWARN.TXT" For Input As #WARN_HANDLE%
If (Err = 0) Then
    Do While Not (EOF(WARN_HANDLE%))
        Line Input #WARN_HANDLE%, h$
        h$ = Trim(h$)
        If (h$ <> "") Then
            ReDim Preserve WarnStr$(AnzWarn%)
            WarnStr$(AnzWarn%) = h$
            AnzWarn% = AnzWarn% + 1
        End If
    Loop
    Close #WARN_HANDLE%
    Kill "WINWARN.TXT"
    
    For i% = 1 To AnzWarn%
        h$ = WarnStr$(i% - 1)
        h2$ = Left$(h$, 16) + vbCrLf + vbCrLf
        h2$ = h2$ + "Lieferant " + Mid$(h$, 18, 6) + vbCrLf
        h3$ = Mid$(h$, 25, 3)
        If (Trim(h3$) <> "") Then
            h2$ = h2$ + h3$ + "% Prognose Mindestumsatz" + vbCrLf
        End If
        h3$ = Mid$(h$, 29, 3)
        If (Trim(h3$) <> "") Then
            h2$ = h2$ + h3$ + "% Prognose Schwellwert" + vbCrLf
        End If
        Call iMsgBox(h2$, vbInformation, "Schwellwert-Automatik")
    Next i%
End If

Call DefErrPop
End Sub

Sub InitSchwellLief(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitSchwellLief")
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

With SchwellLief(ind%)
    .lief = 0
    .rZeit = 0
    .rZeit60 = 0
    .SendungenAlt = 0
    .SendungenPlan = 0
    .AliquotProzent = 0#
    .MindAliquotProzent = 0#
    .AliquotMindestUmsatz = 0#
    .OrgBeobachtSendungen = 0
    .BeobachtSendungen = 0
    .PrognoseRabatt = 0#
    .PrognoseSchwellwert = 0#
    .MindestUmsatz = 0#
    .TabTyp = 0
    
    For i% = 0 To 1
        .OrgBeobachtUmsatz(i%) = 0#
        .BeobachtUmsatz(i%) = 0#
        .UmsatzProSendung(i%) = 0#
        .UmsatzBisher(i%) = 0#
        .UmsatzMitZuordnungen(i%) = 0#
        .UmsatzMitMindest(i%) = 0#
        .UmsatzMitBestLief(i%) = 0#
        .PrognoseUmsatz(i%) = 0#
    Next i%
    
    For i% = 0 To 7
        .SendungenProTag(i%) = 0
    Next i%
End With

Call DefErrPop
End Sub

