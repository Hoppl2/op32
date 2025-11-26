Attribute VB_Name = "modPBA"
Option Explicit

Private Const DefErrModul = "WINPBA.BAS"

Type PbaStruct
  GesKundenProStunde(24) As Long
  KundenProStunde(24) As Double
  PersonalProStunde(24) As Double
  iPersonalProStunde(24) As Integer
  iStundenRest(24) As Integer
  PersonalProTag As Double
  iPersonalProTag As Integer
  KundenProTag As Double
  PersonalWochenStunden As Integer
  StartDatum As Integer
  StopDatum As Integer
End Type

Public PbaDiagrammTyp%(4)
Public PbaDiagrammWas%(4)

Public PbaRec(7) As PbaStruct
Public PbaRec2(7)         As PbaStruct

Public TagName$(7)

Public gVon%, gBis%
Public gAbOffen%, gBisOffen%

Public PersonalWochenStunden%, OrgPersonalWochenStunden%
Public KundenProTag#
Public KundenProArbeitsStunde#

Public StartDatum%, StopDatum%
Public PbaAnalyseAbbruch%


Dim AnzeigeTyp%
Dim AnzWochen%(7)
Dim LetztStartDatum%, LetztStopDatum%

Dim KundenProWoche#, OrgKundenProWoche#

Dim VerkDatum$, VerkLiRe$, VerkZeit%, VerkUser%, DruckDatum$
Dim PersCode%

Public Vergleich%
Public PbaTest%

Dim MaxSchnitt#

Public OrgStartDatum%, OrgStopDatum%

Public GesamtOffen%
Public OffenTagMinuten%
Dim LetztGesamtOffen%
Dim PlanungsTag%

Dim Analysen$()

Dim PbaLen%

Dim DateiName$
Dim DateiHandle%
Dim DateiLen%

Public PbaWahlModus%


Function PbaInit%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PbaInit%")
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
Dim i%, l%, ANALYSE%, erg%
Dim cl$, Y$

TagName$(1) = "Montag"
TagName$(2) = "Dienstag"
TagName$(3) = "Mittwoch"
TagName$(4) = "Donnerstag"
TagName$(5) = "Freitag"
TagName$(6) = "Samstag"
TagName$(7) = "Sonntag"


Call HoleIniPbaDiagramme


AnzeigeTyp% = 0

LetztStartDatum% = -1
LetztStopDatum% = -1

PbaInit% = OpenDaten

PbaLen% = Len(PbaRec(1))

frmPbaDiagramm.Show '1

Call DefErrPop
End Function

Function OpenDaten%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenDaten%")
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

OpenDaten% = False

DateiName$ = "PSTATIS.DAT"
DateiHandle% = FreeFile
Open DateiName$ For Binary Access Read Shared As DateiHandle%
DateiLen% = 64

Seek DateiHandle%, 64! * DateiLen% + 1
Get DateiHandle%, , p
OrgStartDatum% = CVDat(p.datum)

Seek DateiHandle%, LOF(DateiHandle%) - DateiLen% + 1
Get DateiHandle%, , p
OrgStopDatum% = CVDat(p.datum)

StartDatum% = OrgStartDatum%
StopDatum% = OrgStopDatum%

OpenDaten% = True

Call DefErrPop
End Function

'Function SucheDatum%(stDatum%, asatz, VkMax)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SucheDatum%")
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
'Dim gefunden%
'Dim xLinks&, xRechts&, xMitte&
'Dim hStr2$, a$, SuchTagC$
'
'gefunden% = False
'
'a$ = MKI$(stDatum%)
'SuchTagC$ = Mid$(a$, 2, 1) + Mid$(a$, 1, 1)
'
'asatz = 1
'If (SuchTagC$ <> "") Then
'    xLinks = 64: xRechts = VkMax
'
'    While xLinks <= xRechts
'        xMitte = Int((xLinks + xRechts) / 2)
'        Get #DateiHandle%, (xMitte * DateiLen%) + 1, p
'        If (p.datum < SuchTagC$) Then xLinks = xMitte + 1 Else xRechts = xMitte - 1
'    Wend
'    asatz = xMitte: If asatz < 1 Then asatz = 1
'    Get #DateiHandle%, (asatz * DateiLen%) + 1, p
'    While (p.datum < SuchTagC$) And asatz <= VkMax And Not gefunden%
'        Get #DateiHandle%, (asatz * DateiLen%) + 1, p
'        If (p.datum = SuchTagC$) Then gefunden% = -1
'        If (p.datum < SuchTagC$) And Not gefunden% Then asatz = asatz + 1
'    Wend
'    If (asatz > VkMax) Then
'        asatz = VkMax
'    End If
'    If (p.datum = SuchTagC$) Then gefunden% = -1
'End If
'
'SucheDatum% = gefunden%
'
'Call DefErrPop
'End Function
        
Function SucheDatum%(stDatum%, asatz, VkMax)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheDatum%")
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
Dim gefunden%
Dim xLinks&, xRechts&, xMitte&
Dim hStr2$, a$, SuchTagC$

gefunden% = False

a$ = MKI$(stDatum%)
SuchTagC$ = CVDatum(Mid$(a$, 2, 1) + Mid$(a$, 1, 1))

asatz = 1
If (SuchTagC$ <> "") Then
    xLinks = 64: xRechts = VkMax

    While xLinks <= xRechts
        xMitte = Int((xLinks + xRechts) / 2)
        Get #DateiHandle%, (xMitte * DateiLen%) + 1, p
        If (CVDatum(p.datum) < SuchTagC$) Then xLinks = xMitte + 1 Else xRechts = xMitte - 1
    Wend
    asatz = xMitte: If asatz < 1 Then asatz = 1
    Get #DateiHandle%, (asatz * DateiLen%) + 1, p
    While (CVDatum(p.datum) < SuchTagC$) And asatz <= VkMax And Not gefunden%
        Get #DateiHandle%, (asatz * DateiLen%) + 1, p
        If (CVDatum(p.datum) = SuchTagC$) Then gefunden% = -1
        If (CVDatum(p.datum) < SuchTagC$) And Not gefunden% Then asatz = asatz + 1
    Wend
    If (asatz > VkMax) Then
        asatz = VkMax
    End If
    If (CVDatum(p.datum) = SuchTagC$) Then gefunden% = -1
End If

SucheDatum% = gefunden%

Call DefErrPop
End Function
        
Function MachAuswertung%(StartDatum%, StopDatum%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachAuswertung%")
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
Dim i%, j%, k%, MaxRest%, gefunden%, ok%, iVerkDatum%, WoTag%, Stunde%, geoeffnet%
Dim AbOffen%, BisOffen%, AbToleranz%, AbStunde%, AbMinuten%, BisToleranz%, BisStunde%, BisMinuten%
Dim von%, Bis%, iGesBesetzung%, iPersonalProTag%, iStundenPers%, iStundenRest%, Tag%
Dim KuEnde%
Dim asatz&, anz&, anz2&, OrgAnz2&, anz3&
Dim VkMax!, StartZeit!, dauer!, GesamtDauer!, RestDauer!, Prozent!
Dim Summe#, faktor#, dGesBesetzung#, PersonalProTag#, KundenProStunde#, PersonalProStunde#
Dim AltVerkDatum$, hStr2$, a$, h$, DatStr$

MachAuswertung% = False

If ((PbaTest% = 0) And (PersonalWochenStunden% = 0)) Or (GesamtOffen% = 0) Then
    Call DefErrPop: Exit Function
End If

For i% = 0 To 7
    With PbaRec(i%)
        For j% = 0 To 24
            .GesKundenProStunde(j%) = 0
            .KundenProStunde(j%) = 0
            .PersonalProStunde(j%) = 0
            .iPersonalProStunde(j%) = 0
            .iStundenRest(j%) = 0
        Next j%
        .PersonalProTag = 0
        .iPersonalProTag = 0
        .KundenProTag = 0
        .PersonalWochenStunden = 0
        .StartDatum = 0
        .StopDatum = 0
    End With
Next i%


'If (StartDatum% <> LetztStartDatum%) Or (StopDatum% <> LetztStopDatum%) Or (GesamtOffen% <> LetztGesamtOffen%) Then
'    Call FussZeile("[ESC] Abbruch", 0)
'
    VerkDatum$ = Space$(6)
    AltVerkDatum$ = "123456"
    
    For i% = 1 To 7
        AnzWochen%(i%) = 0
    Next i%
    
    For i% = 0 To 23
        For j% = 1 To 7
            PbaRec(j%).GesKundenProStunde(i%) = 0
        Next j%
    Next i%
    
    
    If (DateiHandle% > 0) Then
        
        Seek DateiHandle%, 1
        hStr2$ = Space$(DateiLen%)
        Get DateiHandle%, , hStr2$
        VkMax = CVS(Left(hStr2$, 4)) - 1

        gefunden% = SucheDatum%(StartDatum%, asatz, VkMax)
        If (gefunden% = 0) Then
            gefunden% = SucheDatum%(StartDatum% + 1, asatz, VkMax)
            If (gefunden%) Then
                StartDatum% = StartDatum% + 1
            End If
        End If
        
        If (gefunden%) Then
            frmPbaFortschritt.Show
        Else
            Call DefErrPop: Exit Function
        End If
    
        Seek DateiHandle%, (asatz * DateiLen%) + 1
        
        PbaAnalyseAbbruch% = False
        anz& = 0
        anz2& = 0
        frmPbaFortschritt!prgBestVors.max = VkMax
        StartZeit! = Timer
        
        Do While (1)

            If (anz& Mod 1000 = 0) Then
                frmPbaFortschritt!lblFortschrittStatusWert(0).Caption = anz&
                frmPbaFortschritt!lblFortschrittStatusWert(1).Caption = VerkDatum$
                dauer! = Timer - StartZeit!
                frmPbaFortschritt!lblFortschrittDauerWert(0).Caption = Format$(dauer! \ 60, "##0") + ":" + Format$(dauer! Mod 60, "00")
                Prozent! = (anz& / VkMax) * 100!
                If (Prozent! > 0) Then
                    GesamtDauer! = (dauer! / Prozent!) * 100!
                Else
                    GesamtDauer! = dauer!
                End If
                RestDauer! = GesamtDauer! - dauer!
                frmPbaFortschritt!lblFortschrittDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
'                frmPbaFortschritt!prgBestVors.Value = anz&
                frmPbaFortschritt!lblProzent.Caption = Format$(Prozent!, "##0") + " %"
                
                h$ = Format$(Prozent!, "##0") + " %"
                With frmPbaFortschritt!picProgress
                    .Cls
                    .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
                    .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
                    frmPbaFortschritt!picProgress.Print h$
                    frmPbaFortschritt!picProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
    '                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
                End With
                
                DoEvents
                If (PbaAnalyseAbbruch% = True) Then
                    Exit Do
                End If
            End If
        
            If (EOF(DateiHandle%)) Then
                Exit Do
            End If
        
            hStr2$ = Space$(DateiLen%)
            Get DateiHandle%, , hStr2$
            
            anz& = anz& + 1
            KuEnde% = CVI(Mid$(hStr2$, 23, 2))
            If (KuEnde%) Then
            
                ok% = False
                
                
                DatStr$ = Mid$(hStr2$, 2, 2)
                iVerkDatum% = CVDat(DatStr$)
            
                If (iVerkDatum% >= StartDatum%) And (iVerkDatum% <= StopDatum%) Then
                    VerkDatum$ = sDate(iVerkDatum%)
                    VerkZeit% = CVI(Mid$(hStr2$, 4, 2))
                    ok% = True
                End If
            
                If (iVerkDatum% > StopDatum%) Then
                    Exit Do
                End If
            
                If (ok% = True) Then
                    h$ = Left$(VerkDatum$, 2) + "." + Mid$(VerkDatum$, 3, 2) + "." + Mid$(VerkDatum$, 5, 2)
                    WoTag% = WeekDay(h$, vbMonday)
                    If (WoTag% = 7) Then
                        ok% = False
                    End If
                    If (IstFeiertag(CVDatum(DatStr$))) Then
                        ok% = False
                    End If
                    WoTag% = WoTag% - 1
                End If
            
                If (ok% = True) Then
                    anz2& = anz2& + KuEnde%
                    
                    Stunde% = Int(VerkZeit% \ 100)
                    
                    geoeffnet% = False
                    AbOffen% = OffenRec(WoTag%).iAbOffen
                    BisOffen% = OffenRec(WoTag%).iBisOffen
                    
                    AbToleranz% = AbOffen% - ToleranzOffen%
                    AbStunde% = AbToleranz% \ 100
                    AbMinuten% = AbToleranz% Mod 100
                    If (AbMinuten% > 59) Then
                        AbToleranz% = AbStunde% * 100 + (AbMinuten% - 40)
                    End If
                
                    BisToleranz% = BisOffen% + ToleranzOffen%
                    BisStunde% = BisToleranz% \ 100
                    BisMinuten% = BisToleranz% Mod 100
                    If (BisMinuten% > 59) Then
                        BisToleranz% = (BisStunde% + 1) * 100 + (BisMinuten% - 60)
                    End If
                
                    If (VerkZeit% < AbOffen%) And (VerkZeit% >= AbToleranz%) Then
                        geoeffnet% = True
                        Stunde% = AbOffen% \ 100
                    ElseIf (VerkZeit% > BisOffen%) And (VerkZeit% <= BisToleranz%) Then
                        geoeffnet% = True
                        Stunde% = BisOffen% \ 100
                        If ((BisOffen% Mod 100) = 0) Then
                            Stunde% = Stunde% - 1
                        End If
                    Else
                        For k% = 0 To 1
                            von% = OffenRec(WoTag%).von(k%)
                            Bis% = OffenRec(WoTag%).Bis(k%)
                            If (Bis% > 0) Then
                                If (VerkZeit% = Bis%) And ((VerkZeit% Mod 100) = 0) Then
                                    Stunde% = Stunde% - 1
                                End If
                                If ((VerkZeit% >= von%) And (VerkZeit% <= Bis%)) Then
                                    geoeffnet% = True
                                    Exit For
                                End If
                            End If
                        Next k%
                    End If
                
                    If (geoeffnet%) Then
                        PbaRec(WoTag% + 1).GesKundenProStunde(Stunde%) = PbaRec(WoTag% + 1).GesKundenProStunde(Stunde%) + KuEnde%
                    End If
                End If
            
                If (VerkDatum$ <> AltVerkDatum$) Then
                    AltVerkDatum$ = VerkDatum$
                    If (ok%) Then
                        AnzWochen%(WoTag% + 1) = AnzWochen%(WoTag% + 1) + 1
                    End If
                End If
            End If
        Loop
        
        Unload frmPbaFortschritt
    End If
    
    
    For i% = 1 To 7
        If (AnzWochen%(i%) < 1) Then
            AnzWochen%(i%) = 1
        End If
    Next i%
    
    
    OrgKundenProWoche# = 0#
    MaxSchnitt# = 0#
    For i% = 1 To 7
        Summe# = 0#
        For j% = 0 To 23
            PbaRec(i%).KundenProStunde(j%) = PbaRec(i%).GesKundenProStunde(j%) / AnzWochen%(i%)
            If (PbaRec(i%).KundenProStunde(j%) > MaxSchnitt#) Then
                MaxSchnitt# = PbaRec(i%).KundenProStunde(j%)
            End If
            Summe# = Summe# + PbaRec(i%).KundenProStunde(j%)
        Next j%
        PbaRec(i%).KundenProTag = Summe#
        OrgKundenProWoche# = OrgKundenProWoche# + Summe#
    Next i%
    
    LetztStartDatum% = StartDatum%
    LetztStopDatum% = StopDatum%
    LetztGesamtOffen% = GesamtOffen%
'End If

PbaRec(0).StartDatum = StartDatum%
PbaRec(0).StopDatum = StopDatum%


KundenProWoche# = OrgKundenProWoche#
KundenProTag# = KundenProWoche# / GesamtOffen% * OffenTagMinuten%

If (PbaTest%) Then
Else
    KundenProArbeitsStunde# = KundenProWoche# / PersonalWochenStunden%
End If



iGesBesetzung% = 0
dGesBesetzung# = 0
For j% = 1 To 6
    PersonalProTag# = 0#
    iPersonalProTag% = 0
    For i% = 0 To 23
        KundenProStunde# = PbaRec(j%).KundenProStunde(i%)
        If (KundenProStunde# > 0) Then
            PersonalProStunde# = KundenProStunde# / KundenProArbeitsStunde#
            PbaRec(j%).PersonalProStunde#(i%) = PersonalProStunde#
            
            PersonalProTag# = PersonalProTag# + PersonalProStunde#
            dGesBesetzung# = dGesBesetzung# + PersonalProStunde#
            
            If (PersonalProStunde# < 1#) Then
                PersonalProStunde# = 1#
            End If
            
            iStundenPers% = Int(PersonalProStunde# + 0.501)
            iStundenRest% = Int(PersonalProStunde# * 100# + 0.501) Mod 100
            
            PbaRec(j%).iPersonalProStunde(i%) = iStundenPers%
            PbaRec(j%).iStundenRest(i%) = iStundenRest%
            iPersonalProTag% = iPersonalProTag% + iStundenPers%
            
            iGesBesetzung% = iGesBesetzung% + iStundenPers%
        End If
    Next i%
    PbaRec(j%).PersonalProTag = PersonalProTag#
    PbaRec(j%).iPersonalProTag = iPersonalProTag%
Next j%

If (PbaTest%) Then
    PersonalWochenStunden% = iGesBesetzung%
Else
    While (iGesBesetzung% < PersonalWochenStunden%)
        MaxRest% = 0
        For j% = 1 To 6
            For i% = 0 To 23
                If (PbaRec(j%).iStundenRest(i%) < 50) And (PbaRec(j%).iStundenRest(i%) > MaxRest%) Then
                    Tag% = j%
                    Stunde% = i%
                    MaxRest% = PbaRec(j%).iStundenRest(i%)
                End If
            Next i%
        Next j%
        
        PbaRec(Tag%).iPersonalProStunde(Stunde%) = PbaRec(Tag%).iPersonalProStunde(Stunde%) + 1
        PbaRec(Tag%).iStundenRest(Stunde%) = 0
        
        iGesBesetzung% = iGesBesetzung% + 1
    Wend
    
    While (iGesBesetzung% > PersonalWochenStunden%)
        MaxRest% = 99
        For j% = 1 To 6
            For i% = 0 To 23
                If (PbaRec(j%).iStundenRest(i%) >= 50) And (PbaRec(j%).iStundenRest(i%) < MaxRest%) Then
                    Tag% = j%
                    Stunde% = i%
                    MaxRest% = PbaRec(j%).iStundenRest(i%)
                End If
            Next i%
        Next j%
        
        PbaRec(Tag%).iPersonalProStunde(Stunde%) = PbaRec(Tag%).iPersonalProStunde(Stunde%) - 1
        PbaRec(Tag%).iStundenRest(Stunde%) = 0
        
        iGesBesetzung% = iGesBesetzung% - 1
    Wend
End If

For j% = 1 To 6
    iPersonalProTag% = 0
    For i% = 0 To 23
        iPersonalProTag% = iPersonalProTag% + PbaRec(j%).iPersonalProStunde(i%)
    Next i%
    PbaRec(j%).iPersonalProTag = iPersonalProTag%
Next j%

MachAuswertung% = True

Call DefErrPop
End Function

Sub HoleAnalyse(AnalysenName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleAnalyse")
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
Dim i%, ANALYSE%
Dim Y$

ANALYSE% = FreeFile
If (Dir(AnalysenName + ".pba") <> "") Then
    Open AnalysenName + ".pba" For Binary Access Read As ANALYSE%
    Y$ = Space$(PbaLen%)
    Get ANALYSE%, , Y$
    
'    If (StartDatum% <> CVI(Mid$(Y$, 1, 2))) Or (StopDatum% <> CVI(Mid$(Y$, 3, 2))) Then
'    End If
    
    StartDatum% = CVI(Mid$(Y$, 1, 2))
    StopDatum% = CVI(Mid$(Y$, 3, 2))
    PersonalWochenStunden% = CVI(Mid$(Y$, 5, 2))
    KundenProWoche# = CVD(Mid$(Y$, 7, 8))
    KundenProTag# = CVD(Mid$(Y$, 15, 8))
    KundenProArbeitsStunde# = CVD(Mid$(Y$, 23, 8))
    MaxSchnitt# = CVD(Mid$(Y$, 31, 8))
    
    For i% = 0 To UBound(PbaRec)
        Get ANALYSE%, , PbaRec(i%)
    Next i%
    Close ANALYSE%
    
    LetztStartDatum% = StartDatum%
    LetztStopDatum% = StopDatum%
    
    OrgKundenProWoche# = KundenProWoche#
End If

Call DefErrPop
End Sub

Sub SpeicherAnalyse(AnalysenName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherAnalyse")
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
Dim i%, ANALYSE%
Dim x$

ANALYSE% = FreeFile
Open AnalysenName + ".pba" For Binary Access Write As ANALYSE%
x$ = MKI$(StartDatum%) + MKI$(StopDatum%)
x$ = x$ + MKI$(PersonalWochenStunden%)
x$ = x$ + MKD$(KundenProWoche#)
x$ = x$ + MKD$(KundenProTag#)
x$ = x$ + MKD$(KundenProArbeitsStunde#)
x$ = x$ + MKD$(MaxSchnitt#)
x$ = Left$(x$ + Space$(PbaLen%), PbaLen%)
Put ANALYSE%, , x$

For i% = 0 To UBound(PbaRec)
  Put ANALYSE%, , PbaRec(i%)
Next i%
Close ANALYSE%

Call EinlesenAnalysen

With frmPbaDiagramm.cboAnalysen
    .Enabled = False
    For i% = 0 To (.ListCount - 1)
        If (AnalysenName$ = .List(i%)) Then
            .ListIndex = i%
            Exit For
        End If
    Next i%
    .Enabled = True
End With

Call DefErrPop
End Sub

Sub HoleIniPbaDiagramme()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniPbaDiagramme")
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

Dim i%, ind%, iTyp%, iWas%
Dim l&
Dim h$, key$

For i% = 0 To 4
    h$ = Space$(100)
    key$ = "Diagramm" + Format(i% + 1, "00")
    l& = GetPrivateProfileString("PBA", key$, h$, h$, 101, CurDir + "\winop.ini")
    h$ = Trim$(Left$(h$, l&))
    
    iTyp% = 0
    If (i% = 0) Then
        iWas% = 0
    Else
        iWas% = i% - 1
    End If
    
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        iTyp% = Val(Left$(h$, ind% - 1))
        iWas% = Val(Mid$(h$, ind% + 1))
    End If
    
    PbaDiagrammTyp%(i%) = iTyp%
    PbaDiagrammWas%(i%) = iWas%
Next i%

'h$ = Space$(100)
'Key$ = "Legende"
'l& = GetPrivateProfileString("PVS", Key$, h$, h$, 101, CurDir + "\winop.ini")
'h$ = Trim$(Left$(h$, l&))
'
'LegendenPosStr$ = h$

Call DefErrPop
End Sub

Sub SpeicherIniPbaDiagramme(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniPbaDiagramme")
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
Dim key$, h$

key$ = "Diagramm" + Format(pos% + 1, "00")
h$ = Format(PbaDiagrammTyp%(pos%), "0") + "," + Format(PbaDiagrammWas%(pos%), "0")
l& = WritePrivateProfileString("PBA", key$, h$, CurDir + "\winop.ini")

Call DefErrPop
End Sub

Sub EinlesenAnalysen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenAnalysen")
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
Dim erg%, ind%, i%, k%, gef%, ubou%
Dim SearchHandle&
Dim h$, DirMask$, EntryName$, MinTag$
Dim FindDataRec As WIN32_FIND_DATA

ReDim Analysen$(0)
'Analysen$(0) = "Neue Analyse"

DirMask$ = "*.pba"

SearchHandle& = FindFirstFile(DirMask$, FindDataRec)
If (SearchHandle& <> INVALID_HANDLE_VALUE) Then
    Do
        h$ = FindDataRec.cFileName
        ind% = InStr(h$, Chr$(0))
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        EntryName$ = h$
        
        If ((EntryName$ = ".") Or (EntryName$ = "..")) Then
        ElseIf (FindDataRec.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        Else
            ind% = InStr(EntryName$, ".")
            If (ind% > 0) Then
                EntryName$ = Left$(EntryName$, ind% - 1)
            End If
            ubou% = UBound(Analysen$) + 1
            ReDim Preserve Analysen$(ubou%)
            Analysen$(ubou%) = EntryName$
        End If
        
        erg% = FindNextFile(SearchHandle, FindDataRec)
        If (erg% = 0) Then Exit Do
    Loop
End If
erg% = FindClose(SearchHandle&)

With frmPbaDiagramm.cboAnalysen
    .Enabled = False
    .Clear
    For i% = 1 To UBound(Analysen$)
        .AddItem Analysen$(i%)
    Next i%
    If (.ListCount > 0) Then
        .ListIndex = 0
    End If
    .Enabled = True
End With

Call DefErrPop
End Sub



