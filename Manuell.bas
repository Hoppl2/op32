Attribute VB_Name = "modManuell"
Option Explicit

Public ManuellErg%
Public ManuellBm%, ManuellNm%, ManuellAbholNr%
Public ManuellPzn$, ManuellTxt$
Public ManuellAsatz%, ManuellSsatz%
Public ManuellLief%
'Dim ManuellLaufNr&

Private Const DefErrModul = "MANUELL.BAS"

'Sub ManuellErfassen(pzn$, txt$, MitAnzeige%, Optional TaxBm% = -999)
Sub ManuellErfassen(pzn$, txt$, MitAnzeige%, Optional bm% = 1, Optional nm% = 0, Optional kiste% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ManuellErfassen")
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
ManuellPzn$ = pzn$
ManuellTxt$ = txt$
ManuellBm% = bm%
ManuellNm% = nm%
ManuellAbholNr% = kiste%

Call ManuellBefuellen

If (MitAnzeige%) Then
    Load frmErfassung
    With frmErfassung
        .txtManuell(0).text = ManuellBm%
        .txtManuell(1).text = ManuellNm%
        .Show 1
    End With
    If (ManuellErg%) Then
        Call ActProgram.ManuellSpeichern(True)
    End If
Else
'    ManuellBm% = -(TaxBm%)
    Call ActProgram.ManuellSpeichern
End If

Call DefErrPop
End Sub

Sub ManuellBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ManuellBefuellen")
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
ManuellAsatz% = 0
FabsErrf% = ast.IndexSearch(0, ManuellPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ManuellAsatz% = CInt(FabsRecno&)
    ast.GetRecord (FabsRecno& + 1)
End If

If (ManuellAsatz% = 0) Then
    Call Taxe2ast(ManuellPzn$)
End If
        
ManuellSsatz% = 0
'ManuellBm% = 1
FabsErrf% = ass.IndexSearch(0, ManuellPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ManuellSsatz% = CInt(FabsRecno&)
    ass.GetRecord (FabsRecno& + 1)
'    ManuellBm% = ass.Bm
'    If (ass.vBm > 0) Then ManuellBm% = ass.vBm
End If

'ManuellNm% = 0

Call DefErrPop
End Sub

Sub NeuerDS(Optional ActWu$ = "", Optional RetourZusatz$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NeuerDS")
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
Dim i%, km%, bkMax%, erg%, x1%, xheute%, ind%
Dim aep#, AVP#
Dim txt$, xpzn$, AbCode$, wg$, X$, h$
Dim PreisFaktor

If (ManuellPzn$ = "9999999") Then
    txt$ = ManuellTxt$
    Call CharToOem(txt$, txt$)
    wg$ = "9"
    AbCode$ = "A"
    aep# = 0#
    AVP# = 0#
Else
    txt$ = ast.kurz + ast.meng + ast.meh
    
    wg$ = Left$(ast.wg, 1)
    AbCode$ = Mid$(" AX", InStr("A ", ast.abl) + 1, 1)
            
    aep# = ast.aep
    AVP# = ast.AVP
            
    If (para.Land = "A") And (wg$ = "3") Then
        Mid$(txt$, 29, 5) = ast.herst
        If (Val(ast.herst) <> 0) And (Val(ast.meng) <> 0) Then
            PreisFaktor = Val(ast.herst) / Val(ast.meng)
            aep# = aep# * PreisFaktor
            AVP# = AVP# * PreisFaktor
        End If
    '    If val(ast.PZN2$) = 0 Or val(ast.Meng2$) = 0 Then GoSub Einwieger
    '    If val(ast.PZN2$) = 0 Then apzn$ = "9999999"
    End If
End If

        
X$ = txt$
X$ = LTrim$(X$)
X$ = RTrim$(X$)
If (ManuellPzn$ = "9999999") Then
    For i% = 1 To Len(X$)
        Mid$(X$, i%, 1) = UCase$(Mid$(X$, i%, 1))
    Next i%
End If
txt$ = Left$(X$ + Space$(35), 35)
      

km% = Abs(ManuellBm%)
ww.pzn = ManuellPzn$
ww.txt = txt$

ww.Lief = 0
If (ManuellLief% > 0) Then
    ww.Lief = ManuellLief%
End If
If (ManuellSsatz% > 0) Then
    If (ass.Lief > 0) Then
        ww.Lief = ass.Lief
    End If
End If

ww.bm = ManuellBm%
ww.asatz = ManuellAsatz%
ww.ssatz = ManuellSsatz%
ww.best = " "
ww.nm = ManuellNm%
ww.aep = aep#
ww.abl = AbCode$
ww.wg = wg$
ww.AVP = AVP#
ww.km = km%
ww.absage = 0
ww.angebot = 0
ww.auto = Chr$(0)

ww.alt = Chr$(0)
If (ManuellSsatz% > 0) Then
    x1% = ass.lld
    h$ = Format(Now, "DDMMYY")
    xheute% = iDate(h$)
    If ((x1% + para.MonNBest) < xheute%) Then ww.alt = "?"
End If


ww.nnart = 0
ww.NNAEP = 0#
ww.besorger = " "
ww.AbholNr = 0

ww.aktivlief = 0
ww.aktivind = 0

'ww.beklaufnr = Val(Format(Day(Date), "00") + Right$(Format(Now, "HHMMSS"), 4) + Right$(ManuellPzn$, 3))
ww.BekLaufNr = CalcLaufNr&(ww.pzn)
AnzeigeLaufNr& = ww.BekLaufNr
    

ww.zugeordnet = Chr$(0)
ww.zukontrollieren = Chr$(0)
ww.fixiert = Chr$(0)
ww.DirektTyp = 0

For i% = 0 To 5
    ww.actkontrolle(i%) = 111
Next i%
ww.actzuordnung = 111
ww.PosLag = 0

ww.herst = ast.herst

'If (ast.Rez = "SG") Then
'    ww.IstBtm = 1
'Else
'    ww.IstBtm = 0
'End If

ww.loesch = 0
ww.status = 1
      
ww.IstAltLast = 0
ww.WuStatus = 0
ww.LmStatus = 0
ww.RmStatus = 0
ww.LmAnzGebucht = 0
ww.RmAnzGebucht = 0

ww.IstSchwellArtikel = 0
ww.OrgZeit = 0

ww.zr = -123.45!
                            

If (ActWu$ <> "") Then
    ww.Lief = Asc(Left$(ActWu$, 1))
    ww.WuBestDatum = Mid$(ActWu$, 2, 6)
    ww.WuBestZeit = MKI(Mid$(ActWu$, 8, 4))

    ww.status = 2
    
    ww.WuAEP = ww.aep
    ww.WuBm = Abs(ManuellBm%)
    ww.WuNm = Abs(ManuellNm%)
    
    ww.WuStat = "J"
    If (ManuellBm% < 0) Or (ManuellNm% < 0) Then ww.WuStat = "N"
    
    ww.WuRm = Abs(ManuellBm%)
    ww.WuLm = Abs(ManuellBm%) + Abs(ManuellNm%)
    ww.WuAm = 0
    ww.WuAblDatum = Space$(6)
    ww.WuLa = 0     'mm%    ????
    ww.WuAVP = ww.AVP
    ww.WuBelegDatum = 0
    ww.WuBeleg = Space$(10)
    ww.WuRetMenge = 0
'    ww.WuAnzBereitsGebucht = 0
'    ww.WuFertig = 0
    ww.WuNNaepOk = 0
    ww.WuNNart = 0
    ww.WuNNAep = 0
    ww.WuText = Space$(Len(ww.WuText))
    ww.WuNeuLm = 0
    ww.WuNeuRm = 0
    ww.WuNeuZiel = 0

    ww.aktivlief = 0
    ww.aktivind = 0
    
    If (ManuellAbholNr% > 0) Then
        ww.AbholNr = ManuellAbholNr%
        ww.auto = "v"
    End If


    If (RetourZusatz$ <> "") Then
'        ww.WuFertig = 3
    
        ind% = InStr(RetourZusatz$, "@")
        h$ = Left$(RetourZusatz$, ind% - 1)
        RetourZusatz$ = Mid$(RetourZusatz$, ind% + 1)
        ww.WuBeleg = h$
    
        ind% = InStr(RetourZusatz$, "@")
        h$ = Left$(RetourZusatz$, ind% - 1)
        RetourZusatz$ = Mid$(RetourZusatz$, ind% + 1)
        ww.WuBelegDatum = iDate(h$)
    
        ind% = InStr(RetourZusatz$, "@")
        h$ = Left$(RetourZusatz$, ind% - 1)
        RetourZusatz$ = Mid$(RetourZusatz$, ind% + 1)
        ww.bm = 0
        ww.nm = 0
        ww.WuNeuLm = -Val(h$)
        ww.WuNeuRm = -Val(h$)
    
        ww.WuBm = 0
        ww.WuNm = 0
        ww.WuRm = Val(h$)
        ww.WuLm = Val(h$)
        
        ind% = InStr(RetourZusatz$, "@")
        h$ = Left$(RetourZusatz$, ind% - 1)
        RetourZusatz$ = Mid$(RetourZusatz$, ind% + 1)
        If (Trim$(h$) <> "") Then ww.WuAblDatum = "01" + h$
        ww.WuAEP = CDbl(RetourZusatz$)
    
        ww.IstAltLast = 1
    End If
End If


ww.SatzLock (1)
ww.GetRecord (1)
bkMax% = ww.erstmax
bkMax% = bkMax% + 1
ww.erstmax = bkMax%
ww.PutRecord (1)
ww.PutRecord (bkMax% + 1)
ww.SatzUnLock (1)

Call DefErrPop
End Sub

Function ArtikelAuswahl$(Optional MitSpeichern% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ArtikelAuswahl$")
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
Dim ind%, vorz%, bm%, nm%
Dim mErg$, h$, pzn$, txt$, ret$
            
ret$ = ""
ManuellLief% = 0
        
If (LieferantenAbfrage%) Then
    mErg$ = MatchCode(1, pzn$, txt$, False, False)
    If (mErg$ <> "") Then
        ManuellLief% = Val(pzn$)
    End If
End If

mErg$ = MatchCode(0, pzn$, txt$, False, False)
If (mErg$ <> "") Then
    Do
        If (mErg$ = "") Then Exit Do
        
        ind% = InStr(mErg$, vbTab)
        h$ = Left$(mErg$, ind% - 1)
        mErg$ = Mid$(mErg$, ind% + 1)
        
        vorz% = 1
        If (Right$(h$, 1) = "-") Then
            vorz% = -1
            h$ = Left$(h$, Len(h$) - 1)
        End If
        ind% = InStr(h$, "@")
        pzn$ = Left$(h$, ind% - 1)
        h$ = Mid$(h$, ind% + 1)
        
        ret$ = pzn$
        
        ind% = InStr(h$, "@")
        txt$ = Left$(h$, ind% - 1)
        h$ = Mid$(h$, ind% + 1)
        
        ind% = InStr(h$, "@")
        bm% = Val(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
        
        ind% = InStr(h$, "@")
        nm% = Val(Left$(h$, ind% - 1))
        h$ = Mid$(h$, ind% + 1)
    
        If (MitSpeichern%) Then Call ManuellErfassen(pzn$, txt$, False, bm%, nm%)
    Loop
End If

ManuellLief% = 0

ArtikelAuswahl$ = ret$

Call DefErrPop
End Function

Function CalcLaufNr&(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcLaufNr&")
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
Static iLaufNr%

If (iLaufNr% = 0) Then
    iLaufNr% = Val(Right$(Format(Day(Date), "00"), 1) + Right$(Format(Now, "HHMMSS"), 2))
End If

iLaufNr% = iLaufNr% + 1
If (iLaufNr% > 999) Then
    iLaufNr% = 1
End If

CalcLaufNr& = Val(Format(iLaufNr%, "000") + Left$(pzn$, 6))

Call DefErrPop
End Function


