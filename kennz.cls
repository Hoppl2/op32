VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKennzeichen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


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

Private Const DefErrModul = "KENNZ.CLS"

Private Sub Class_Initialize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Class_Initialize")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
dKurz$ = "KZ"
dBezeichnung$ = "Kennzeichen"
dName$ = "Kz"

If (Para1.Land = "D") Then

    dAnzFields% = 17
    ReDim fKurz$(dAnzFields% - 1)
    ReDim fBezeichnung$(dAnzFields% - 1)
    'ReDim fStart%(dAnzFields% - 1)
    'ReDim fLaenge%(dAnzFields% - 1)
    'ReDim fTyp$(dAnzFields% - 1)
    'ReDim fKonv%(dAnzFields% - 1)
    
    fKurz$(0) = "FEST"
    fBezeichnung$(0) = "Festbetrag"
    
    fKurz$(1) = "NEG"
    fBezeichnung$(1) = "Negativliste"
    
    fKurz$(2) = "REZ"
    fBezeichnung$(2) = "Rezeptpflicht"
    
    fKurz$(3) = "BTM"
    fBezeichnung$(3) = "Betäubungsmittel"
    
    fKurz$(4) = "RW"
    fBezeichnung$(4) = "Rückruf/Widerruf"
    
    fKurz$(5) = "NAM"
    fBezeichnung$(5) = "Nettoartikel"
    
    fKurz$(6) = "LAC"
    fBezeichnung$(6) = "Lagercode"
    
    fKurz$(7) = "LIC"
    fBezeichnung$(7) = "Liefercode"
    
    fKurz$(8) = "IMP"
    fBezeichnung$(8) = "Importe"
    
    fKurz$(9) = "ALT"
    fBezeichnung$(9) = "Alternativpackungen"
    
    fKurz$(10) = "HILF"
    fBezeichnung$(10) = "Hilfsmittel"
    
    fKurz$(11) = "ABL"
    fBezeichnung$(11) = "Ablaufcode"
    
    fKurz$(12) = "VC"
    fBezeichnung$(12) = "Veränderungscode"
    
    fKurz$(13) = "KZA"
    fBezeichnung$(13) = "Kz-Aufschlag"
    
    fKurz$(14) = "INT"
    fBezeichnung$(14) = "Interne Streichung"
    
    fKurz$(15) = "BSTAT"
    fBezeichnung$(15) = "Artikel-Status"
    
    fKurz$(16) = "SONDERPR"
    fBezeichnung$(16) = "Sonderpreis"
Else

    dAnzFields% = 12
    ReDim fKurz$(dAnzFields% - 1)
    ReDim fBezeichnung$(dAnzFields% - 1)
    'ReDim fStart%(dAnzFields% - 1)
    'ReDim fLaenge%(dAnzFields% - 1)
    'ReDim fTyp$(dAnzFields% - 1)
    'ReDim fKonv%(dAnzFields% - 1)
    
    fKurz$(0) = "REZ"
    fBezeichnung$(0) = "Rezeptpflicht"
    
    fKurz$(1) = "BTM"
    fBezeichnung$(1) = "Betäubungsmittel"
    
    fKurz$(2) = "LAC"
    fBezeichnung$(2) = "Lagercode"
    
    fKurz$(3) = "LIC"
    fBezeichnung$(3) = "Liefercode"
    
    fKurz$(4) = "ABL"
    fBezeichnung$(4) = "Ablaufcode"
    
    fKurz$(5) = "VC"
    fBezeichnung$(5) = "Veränderungscode"
    
    fKurz$(6) = "KZA"
    fBezeichnung$(6) = "Kz-Aufschlag"
    
    fKurz$(7) = "INT"
    fBezeichnung$(7) = "Interne Streichung"
    
    fKurz$(8) = "BSTAT"
    fBezeichnung$(8) = "Artikel-Status"

    fKurz$(9) = "REZZ"
    fBezeichnung$(9) = "Rezeptzeichen"

    fKurz$(10) = "KASZ"
    fBezeichnung$(10) = "Kassenzeichen"
    
    fKurz$(11) = "KASZZU"
    fBezeichnung$(11) = "Kassenzeichen-Zusatz"
End If


Call clsError.DefErrPop
End Sub

Sub FreeClass()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FreeClass")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrPop
End Sub

Function SucheFeldInd%(Typ%, SuchFeld$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SucheFeldInd%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
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

Call clsError.DefErrPop
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

Sub CheckKzDS(pzn$, iKz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckKzDS")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%
Dim s$, h$

If (ArtikelDBok) Then
    If (InStr(UCase(iKz), "SONDERPR") > 0) Then
'        On Error Resume Next
'        SonderPrAdoRec.Close
'        Err.Clear
'        On Error GoTo DefErr
'        SQLStr$ = "SELECT * FROM SonderPreise WHERE PZN = " + pzn$
'        SonderPrAdoRec.Open SQLStr, ArtikelDB1.ActiveConn
'        If (SonderPrAdoRec.EOF = False) Then
'            SonderPrDa% = True
'        End If
        SonderPrDa = CheckSonderPreis(pzn)
    ElseIf (ArtNeu% = False) Then
        On Error Resume Next
        ArtikelRec.Close
        Err.Clear
        On Error GoTo DefErr
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn$
        ArtikelRec.Open SQLStr, ArtikelDB1.ActiveConn
        If (ArtikelRec.EOF = False) Then
            Call ArtikelDB1.DB2DOS(ArtikelRec)
            ArtDa% = True
        End If
        AstDa% = ArtDa%
        AssDa% = ArtDa%
        ArtNeu% = True
        AstNeu% = True
        AssNeu% = True
    End If
    
    Call clsError.DefErrPop: Exit Sub
End If

h$ = iKz$
If (Right$(h$, 1) <> ",") Then h$ = h$ + ","
Do
    If (h$ = "") Then Exit Do
    ind% = InStr(h$, ",")
    s$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)

    s$ = UCase$(s$)
    
    ind% = SucheFeldInd%(0, s$)
    If (ind% = 14) Then
        Call clsDat.CheckDateiDS(pzn$, "ASS", True)
    ElseIf (ind% = 16) Then
        Call clsDat.CheckDateiDS(pzn$, "SONDERPR", True)
    ElseIf (ind% >= 0) Then
        Call clsDat.CheckDateiDS(pzn$, "TAXE", True)
        Call clsDat.CheckDateiDS(pzn$, "AST", True)
    End If
Loop


Call clsError.DefErrPop
End Sub

Function HoleKzValue$(iKz$, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleKzValue$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%, iVal%
Dim s$, h$, ret$, KzVal$

If (TaxeAdoDBok) Then
    HoleKzValue = HoleKzValueADO$(iKz$, pzn$): Call clsError.DefErrPop: Exit Function
End If

ret$ = ""

h$ = iKz$
If (Right$(h$, 1) <> ",") Then h$ = h$ + ","
Do
    If (h$ = "") Then Exit Do
    ind% = InStr(h$, ",")
    s$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)

    s$ = UCase$(s$)
    KzVal$ = " "
    
    ind% = SucheFeldInd%(0, s$)
    
    If (Para1.Land = "D") Then
    
        Select Case ind%
            Case 0  'Festbetrag
                If (AstDa%) Then
                    KzVal$ = Left$(Ast1.kasz, 1)
                ElseIf (TaxeDa%) Then
                    If (TaxeRec!FestKz) Then KzVal$ = "F"
                End If
            Case 1  'Negativliste
                If (AstDa%) Then
                    If (Left$(Ast1.kas, 1) = "*") Then KzVal$ = "*"
                ElseIf (TaxeDa%) Then
                    If (TaxeRec!NegListe) Then KzVal$ = "*"
                End If
            Case 2  'Rezeptpflicht
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "+") Then KzVal$ = "+"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!AbgabeBest
                    If (iVal% = 1) Or (iVal% = 5) Or (iVal% = 6) Then KzVal$ = "+"
                End If
            Case 3  'Betäubungsmittel
                If (AstDa%) Then
                    If (Ast1.rez = "SG") Then KzVal$ = "B"
                ElseIf (TaxeDa%) Then
                    If (TaxeRec!BTMKz) Then KzVal$ = "B"
                End If
            Case 4  'Rückruf/Widerruf
                If (AstDa%) Then
                    If (Left$(Ast1.kas, 2) = "RW") Then KzVal$ = "R"
                ElseIf (TaxeDa%) Then
                    If (TaxeRec!RufKz) Then KzVal$ = "R"
                End If
            Case 5  'Nettoartikel
                If (AstDa%) Then
                    If (InStr(Ast1.rez, "P") > 0) Then KzVal$ = "P"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!AbgabeBest
                    If (iVal% = 2) Or (iVal% = 3) Then KzVal$ = "P"
                End If
            Case 6  'Lagercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lac
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!Lagerung
                    If (iVal% > 0) Then KzVal$ = Format(iVal%, "0")
                End If
            Case 7  'Liefercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lic
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!UeberGh
                    If (iVal% = 1) Then
                        KzVal$ = "D"
                    ElseIf (iVal% = 2) Then
                        KzVal$ = "K"
                    End If
                End If
            Case 8  'Importe
                If (TaxeDa%) Then
                    If (TaxeRec!OriginalPZN > 0) Then
                        If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                            KzVal$ = "i"
                        Else
                            KzVal$ = "I"
                        End If
                    End If
                End If
            Case 9  'Alternativpackungen
                If (TaxeDa%) Then
                    If (TaxeRec!AlternativPackung > 0) Then
                        If (TaxeRec!AlternativPackung = TaxeRec!pzn) Then
                            KzVal$ = "G"
                        Else
                            KzVal$ = "g"
                        End If
                    End If
                End If
            Case 10  'Hilfsmittel
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "A") Then KzVal$ = "H"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!Warengruppe
                    If ((iVal% >= 10) And (iVal% <> 55)) Then KzVal$ = "H"
                End If
            Case 11  'Ablaufcode
                If (AstDa%) Then
                    KzVal$ = Ast1.abl
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!VerfallArt
                    If (iVal% = 1) Then KzVal$ = "A"
                End If
            Case 12  'Veränderungscode
                If (AstDa%) Then
                    KzVal$ = Ast1.vc
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!ArtStatus
                End If
            Case 13  'Kennzeichen-Aufschlag
                If (AstDa%) Then
                    KzVal$ = Ast1.ka
                End If
            Case 14  'Interne Streichung
                If (AssDa%) Then
                    If (Ass1.halt = "S") Then KzVal$ = "s"
                End If
            Case 15  'Bestellstatus
                KzVal$ = clsDialog.TesteBstatus$(pzn$)
            Case 16  'Sonderpreis
                If (SonderPrDa%) Then
                    If (Not (IsNull(SonderPrRec!Gültigab))) Then
                        If (Year(SonderPrRec!Gültigab) > 2000) Then
                            KzVal$ = KzVal$ + Format(SonderPrRec!Gültigab, "DD.MM.YY")
                        End If
                    End If
                    If (Not (IsNull(SonderPrRec!GültigBis))) Then
                        If (Year(SonderPrRec!GültigBis) > 2000) Then
                            KzVal$ = KzVal$ + "-" + Format(SonderPrRec!GültigBis, "DD.MM.YY")
                        End If
                    End If
                    KzVal$ = KzVal$ + ": " + Format(SonderPrRec!Sonderpreis / 100!, "0.00")
                End If
        End Select
        
    Else
        
        Select Case ind%
    
            Case 0  'Rezeptpflicht
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "+") Then KzVal$ = "+"
                ElseIf (TaxeDa%) Then
                    If (Left$(TaxeRec!RezA, 1) = "+") Then KzVal$ = "+"
                End If
            Case 1  'Betäubungsmittel
                If (AstDa%) Then
                    If (Ast1.rez = "SG") Then KzVal$ = "B"
                ElseIf (TaxeDa%) Then
                    If (TaxeRec!BTMKz) Then KzVal$ = "B"
                End If
            Case 2  'Lagercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lac
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!LagerungA
                End If
            Case 3  'Liefercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lic
                ElseIf (TaxeDa%) Then
                End If
            Case 4  'Ablaufcode
                If (AstDa%) Then
                    KzVal$ = Ast1.abl
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeRec!VerfallArt
                    If (iVal% = 1) Then KzVal$ = "A"
                End If
            Case 5  'Veränderungscode
                If (AstDa%) Then
                    KzVal$ = Ast1.vc
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!ArtStatus
                End If
            Case 6  'Kennzeichen-Aufschlag
                If (AstDa%) Then
                    KzVal$ = Ast1.ka
                End If
            Case 7  'Interne Streichung
                If (AssDa%) Then
                    If (Ass1.halt = "S") Then KzVal$ = "s"
                End If
            Case 8  'Bestellstatus
                KzVal$ = clsDialog.TesteBstatus$(pzn$)
            Case 9  'Rezeptzeichen
                If (AstDa%) Then
                    KzVal$ = Ast1.rez
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!RezA
                End If
            Case 10  'Kassenzeichen
                If (AstDa%) Then
                    KzVal$ = Ast1.kas
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!DarForm
                End If
            Case 11  'Kassenzeichen-Zusatz
                If (AstDa%) Then
                    KzVal$ = Ast1.kasz
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeRec!KasZuA
                End If
        End Select
    
    End If
    
    ret$ = ret$ + KzVal$
    
Loop

HoleKzValue$ = ret$

Call clsError.DefErrPop
End Function

Function HoleKzValueADO$(iKz$, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleKzValueADO$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%, iVal%
Dim s$, h$, ret$, KzVal$

ret$ = ""

h$ = iKz$
If (Right$(h$, 1) <> ",") Then h$ = h$ + ","
Do
    If (h$ = "") Then Exit Do
    ind% = InStr(h$, ",")
    s$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)

    s$ = UCase$(s$)
    KzVal$ = " "
    
    ind% = SucheFeldInd%(0, s$)
    
    If (Para1.Land = "D") Then
    
        Select Case ind%
            Case 0  'Festbetrag
                If (AstDa%) Then
                    KzVal$ = Left$(Ast1.kasz, 1)
                ElseIf (TaxeDa%) Then
                    If (TaxeAdoRec!FestKz) Then KzVal$ = "F"
                End If
            Case 1  'Negativliste
                If (AstDa%) Then
                    If (Left$(Ast1.kas, 1) = "*") Then KzVal$ = "*"
                ElseIf (TaxeDa%) Then
                    If (TaxeAdoRec!NegListe) Then KzVal$ = "*"
                End If
            Case 2  'Rezeptpflicht
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "+") Then KzVal$ = "+"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!AbgabeBest
                    If (iVal% = 1) Or (iVal% = 5) Or (iVal% = 6) Then KzVal$ = "+"
                End If
            Case 3  'Betäubungsmittel
                If (AstDa%) Then
                    If (Ast1.rez = "SG") Then KzVal$ = "B"
                ElseIf (TaxeDa%) Then
                    If (TaxeAdoRec!BTMKz) Then KzVal$ = "B"
                End If
            Case 4  'Rückruf/Widerruf
                If (AstDa%) Then
                    If (Left$(Ast1.kas, 2) = "RW") Then KzVal$ = "R"
                ElseIf (TaxeDa%) Then
                    If (TaxeAdoRec!RufKz) Then KzVal$ = "R"
                End If
            Case 5  'Nettoartikel
                If (AstDa%) Then
                    If (InStr(Ast1.rez, "P") > 0) Then KzVal$ = "P"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!AbgabeBest
                    If (iVal% = 2) Or (iVal% = 3) Then KzVal$ = "P"
                End If
            Case 6  'Lagercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lac
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!Lagerung
                    If (iVal% > 0) Then KzVal$ = Format(iVal%, "0")
                End If
            Case 7  'Liefercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lic
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!UeberGh
                    If (iVal% = 1) Then
                        KzVal$ = "D"
                    ElseIf (iVal% = 2) Then
                        KzVal$ = "K"
                    End If
                End If
            Case 8  'Importe
                If (TaxeDa%) Then
                    If (TaxeAdoRec!OriginalPZN > 0) Then
                        If (TaxeAdoRec!OriginalPZN = TaxeAdoRec!pzn) Then
                            KzVal$ = "i"
                        Else
                            KzVal$ = "I"
                        End If
                    End If
                End If
            Case 9  'Alternativpackungen
                If (TaxeDa%) Then
                    If (TaxeAdoRec!AlternativPackung > 0) Then
                        If (TaxeAdoRec!AlternativPackung = TaxeAdoRec!pzn) Then
                            KzVal$ = "G"
                        Else
                            KzVal$ = "g"
                        End If
                    End If
                End If
            Case 10  'Hilfsmittel
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "A") Then KzVal$ = "H"
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!Warengruppe
                    If ((iVal% >= 10) And (iVal% <> 55)) Then KzVal$ = "H"
                End If
            Case 11  'Ablaufcode
                If (AstDa%) Then
                    KzVal$ = Ast1.abl
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!VerfallArt
                    If (iVal% = 1) Then KzVal$ = "A"
                End If
            Case 12  'Veränderungscode
                If (AstDa%) Then
                    KzVal$ = Ast1.vc
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!ArtStatus
                End If
            Case 13  'Kennzeichen-Aufschlag
                If (AstDa%) Then
                    KzVal$ = Ast1.ka
                End If
            Case 14  'Interne Streichung
                If (AssDa%) Then
                    If (Ass1.halt = "S") Then KzVal$ = "s"
                End If
            Case 15  'Bestellstatus
                KzVal$ = clsDialog.TesteBstatus$(pzn$)
            Case 16  'Sonderpreis
                If (SonderPrDa%) Then
                    If (Not (IsNull(SonderPrAdoRec!Gültigab))) Then
                        If (Year(SonderPrAdoRec!Gültigab) > 2000) Then
                            KzVal$ = KzVal$ + Format(SonderPrAdoRec!Gültigab, "DD.MM.YY")
                        End If
                    End If
                    If (Not (IsNull(SonderPrAdoRec!GültigBis))) Then
                        If (Year(SonderPrAdoRec!GültigBis) > 2000) Then
                            KzVal$ = KzVal$ + "-" + Format(SonderPrAdoRec!GültigBis, "DD.MM.YY")
                        End If
                    End If
                    KzVal$ = KzVal$ + ": " + Format(SonderPrAdoRec!Sonderpreis, "0.00")
                End If
        End Select
        
    Else
        
        Select Case ind%
    
            Case 0  'Rezeptpflicht
                If (AstDa%) Then
                    If (Left$(Ast1.rez, 1) = "+") Then KzVal$ = "+"
                ElseIf (TaxeDa%) Then
                    If (Left$(TaxeAdoRec!RezA, 1) = "+") Then KzVal$ = "+"
                End If
            Case 1  'Betäubungsmittel
                If (AstDa%) Then
                    If (Ast1.rez = "SG") Then KzVal$ = "B"
                ElseIf (TaxeDa%) Then
                    If (TaxeAdoRec!BTMKz) Then KzVal$ = "B"
                End If
            Case 2  'Lagercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lac
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!LagerungA
                End If
            Case 3  'Liefercode
                If (AstDa%) Then
                    KzVal$ = Ast1.lic
                ElseIf (TaxeDa%) Then
                End If
            Case 4  'Ablaufcode
                If (AstDa%) Then
                    KzVal$ = Ast1.abl
                ElseIf (TaxeDa%) Then
                    iVal% = TaxeAdoRec!VerfallArt
                    If (iVal% = 1) Then KzVal$ = "A"
                End If
            Case 5  'Veränderungscode
                If (AstDa%) Then
                    KzVal$ = Ast1.vc
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!ArtStatus
                End If
            Case 6  'Kennzeichen-Aufschlag
                If (AstDa%) Then
                    KzVal$ = Ast1.ka
                End If
            Case 7  'Interne Streichung
                If (AssDa%) Then
                    If (Ass1.halt = "S") Then KzVal$ = "s"
                End If
            Case 8  'Bestellstatus
                KzVal$ = clsDialog.TesteBstatus$(pzn$)
            Case 9  'Rezeptzeichen
                If (AstDa%) Then
                    KzVal$ = Ast1.rez
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!RezA
                End If
            Case 10  'Kassenzeichen
                If (AstDa%) Then
                    KzVal$ = Ast1.kas
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!DarForm
                End If
            Case 11  'Kassenzeichen-Zusatz
                If (AstDa%) Then
                    KzVal$ = Ast1.kasz
                ElseIf (TaxeDa%) Then
                    KzVal$ = TaxeAdoRec!KasZuA
                End If
        End Select
    
    End If
    
    ret$ = ret$ + KzVal$
    
Loop

HoleKzValueADO$ = ret$

Call clsError.DefErrPop
End Function


Function CheckSonderPreis%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckSonderPreis%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%, ok%
Dim SQLStr$

ret = 0

If (Val(pzn) = 0) Then
    CheckSonderPreis = 0
    Call clsError.DefErrPop: Exit Function
End If

If (ArtikelDBok) Then
    On Error Resume Next
    SonderPrAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    
    SQLStr = "SELECT * FROM Sonderpreise WHERE PZN=" + CStr(pzn) + " AND LöschKZ=0 ORDER BY GültigAb"
    SonderPrAdoRec.Open SQLStr, ArtikelDB1.ActiveConn
    Do While Not (SonderPrAdoRec.EOF)
        If (DateValue(SonderPrAdoRec!Gültigab) <= DateValue(Now) Or Format(SonderPrAdoRec!Gültigab, "ddmmyy") = "010100") And (DateValue(SonderPrAdoRec!GültigBis) >= DateValue(Now) Or Format(SonderPrAdoRec!GültigBis, "ddmmyy") = "010100") And SonderPrAdoRec!Sonderpreis > 0# Then
            ret = True
            Exit Do
        End If
        
        SonderPrAdoRec.MoveNext
    Loop
Else
End If

CheckSonderPreis = ret

Call clsError.DefErrPop
End Function
