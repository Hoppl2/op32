Attribute VB_Name = "modManuell"
Option Explicit

Public ManuellErg%
Public ManuellBm%, ManuellNm%
Public ManuellPzn$, ManuellTxt$
Dim ManuellAsatz%, ManuellSsatz%
'Dim ManuellLaufNr&

Sub ManuellErfassen(pzn$, txt$)

ManuellPzn$ = pzn$
ManuellTxt$ = txt$
Load frmErfassung
Call ManuellBefuellen
frmErfassung.Show 1
If (ManuellErg%) Then
    Call ManuellSpeichern
    Call frmAction.AuslesenBestellung(True, False, True)
End If

End Sub

Sub ManuellBefuellen()

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
ManuellBm% = 1
FabsErrf% = ass.IndexSearch(0, ManuellPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ManuellSsatz% = CInt(FabsRecno&)
    ass.GetRecord (FabsRecno& + 1)
    ManuellBm% = ass.bm
    If (ass.vbm > 0) Then ManuellBm% = ass.vbm
End If

ManuellNm% = 0

With frmErfassung
    .txtManuell(0).text = ManuellBm%
    .txtManuell(1).text = ManuellNm%
End With

End Sub

Sub ManuellSpeichern()
Dim i%, km%, bkMax%, erg%, x1%, xheute%
Dim AEP#, AVP#
Dim txt$, xpzn$, ABCODE$, wg$, X$, h$

If (ManuellSsatz% = 0) Then
    If (ManuellPzn$ = "9999999") Or (ManuellAsatz% = 0) Then
        ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
    Else
        erg% = MsgBox("Ist dieser Artikel bereits auf Lager ?", vbYesNo Or vbInformation)
        If (erg% = vbNo) Then
            ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
        End If
    End If
End If

If (ManuellPzn$ = "9999999") Then
    txt$ = ManuellTxt$
    Call CharToOem(txt$, txt$)
    wg$ = "9"
    ABCODE$ = "A"
    AEP# = 0#
    AVP# = 0#
Else
    txt$ = ast.kurz + ast.meng + ast.meh
    
    wg$ = Left$(ast.wg, 1)
    ABCODE$ = Mid$(" AX", InStr("A ", ast.abl) + 1, 1)
            
    AEP# = ast.AEP
    AVP# = ast.AVP
            
    If (wg$ = "3") Then
    '    If val(ast.PZN2$) = 0 Or val(ast.Meng2$) = 0 Then GoSub Einwieger
    '    Mid$(txt$, 29, 5) = ast.Meng2$
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
bek.pzn = ManuellPzn$
bek.txt = txt$
bek.lief = 0
If (ManuellSsatz% > 0) Then bek.lief = ass.lief

bek.bm = ManuellBm%
bek.asatz = ManuellAsatz%
bek.ssatz = ManuellSsatz%
bek.best = " "
bek.nm = ManuellNm%
bek.AEP = AEP#
bek.abl = ABCODE$
bek.wg = wg$
bek.AVP = AVP#
bek.km = km%
bek.absage = 0
bek.angebot = 0
bek.auto = Chr$(0)

bek.alt = Chr$(0)
If (ManuellSsatz% > 0) Then
    x1% = ass.lld
    h$ = Format(Now, "DDMMYY")
    xheute% = iDate(h$)
    If ((x1% + para.MonNBest) < xheute%) Then bek.alt = "?"
End If


bek.nnart = 0
bek.NNAep = 0#
bek.besorger = " "
bek.AbholNr = 0

bek.aktivlief = 0
bek.aktivind = 0

bek.BekLaufNr = Val(Format(Day(Date), "00") + Right$(Format(Now, "HHMMSS"), 4) + Right$(ManuellPzn$, 3))
AnzeigeLaufNr& = bek.BekLaufNr
    

bek.zugeordnet = Chr$(0)
bek.zukontrollieren = Chr$(0)
bek.fixiert = Chr$(0)
bek.nochzukontrollieren = Chr$(0)

For i% = 0 To 5
    bek.actkontrolle(i%) = 111
Next i%
bek.ActZuordnung = 111
bek.poslag = 0

bek.herst = ast.herst

      
bek.SatzLock (1)
bek.GetRecord (1)
bkMax% = bek.erstmax
bkMax% = bkMax% + 1
bek.erstmax = bkMax%
bek.PutRecord (1)
bek.PutRecord (bkMax% + 1)
bek.SatzUnLock (1)

Call DefErrPop

End Sub

