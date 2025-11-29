Attribute VB_Name = "modTaxieren"
Option Explicit

Public Const MAG_GEFAESS = 0
Public Const MAG_ARBEIT = 1
Public Const MAG_SPEZIALITAET = 2
Public Const MAG_SONSTIGES = 3
Public Const MAG_HILFSTAXE = 4
Public Const MAG_PREISEINGABE = 5
Public Const MAG_ANTEILIG = 6

Public Const MAG_NN = 100

Public Const ABHOLER = "abholer.mdb"
Public Const TAXMUSTER_DB = "taxmustr.mdb"

Type TaxMusterHeader2Struct
    Name As String * 8
    dummy As String * 1
End Type
Type TaxMusterHeaderStruct
    Name As String * 35
    dummy As String * 1
    ActMenge As Integer
    AnzZeilen As Byte
    Inhalt(2) As TaxMusterHeader2Struct
    ErstSatz As Long
    DSnummer As Long
    InStringPos As Byte
    rest As String * 5
End Type
Public TmHeader As TaxMusterHeaderStruct

Type TaxMusterInhaltStruct
    pzn As String * 8
    dummy As String * 1
    kurz As String * 28
    dummy2 As String * 1
    ActMenge As Double
    ActPreis As Double
    flag As Byte
    NextSatz As Long
End Type
Public TmInhalt As TaxMusterInhaltStruct

Type ArbEmbStruct
    kurz As String * 28
    menge As Single
    blank As String * 1
    Meh As String * 2
    kp As Double
    Typ As String * 1
    rest As String * 84
End Type
Public ArbEmbRec As ArbEmbStruct


Type TaxierungStruct
    pzn As String * 8
    kurz As String * 28
    menge As String * 5
    Meh As String * 2
    flag As Byte
    kp As Double
    GStufe As Double
    ActMenge As Double
    ActPreis As Double
    Verwurf As Byte
End Type
Public TaxierRec As TaxierungStruct


'Type AnfMagHeaderStruct
'    VonWann As Long
'    Kiste As Integer
'    WasTun As String * 1
'    RezeptNr As Integer
'    LaufNr As Byte
'End Type
'Public AnfMagHeader As AnfMagHeaderStruct

'Type AnfMagInhaltStruct
'    pzn As String * 7
'    blank As String * 1
'    kurz As String * 28
'    blank2 As String * 1
'    fMenge As Single
'    blank3 As String * 2
'    meh As String * 2
'    blank4 As String * 1
'    aep As Double
'    blank5 As String * 1
'    Kp As Double
'    blank6 As String * 1
'    ActMenge As Double
'    blank7 As String * 1
'    ActPreis As Double
'    blank8 As String * 1
'    Mw As Byte
'    Minans As Byte
'    Gstufe As String * 8
'    blank9 As String * 1
'    Match As String * 38
'    blank10 As String * 1
'    DSnummer As Long
'    flag As Byte
'End Type
'Public AnfMagInhalt As AnfMagInhaltStruct



Public TM_NAMEN%, TM_DATEN%
Public MAG_SPEICHER%

Public TaxmusterDB As Database
Public TaxmusterRec As Recordset
Public TaxmusterDBok%

'Public AbholerDB As Database
'Public AnfMagRec As Recordset


Public TmMengenFaktor#

Public ARBEMB%
Public AnfArb%, AnfEmb%

Public MalFaktor#, NeuMalFaktor#

Public SumPreis#, TeilPreis#, ProzPreis#, TeilMenge#, SollMenge#
Public SumPreisZuz#, TeilPreisZuz#, ProzPreisZuz#
Public SollTaxierTyp%
Public SollPzn$

Public TaxMusterModus%
Public TaxMusterSuch$

Public MagSpeicherIndex%
Public AnfMagIndex&

Public AnfMagPreis#

Public RezRbhLauer%
'Public RezRbhLauerTemp$
Public RezRbhLauerRec As New ADODB.Recordset

Public UnverarbeiteteAbgabe%

Public CannabisExtraktDichte#

Public FiveRxPzns$(50)
Public MaxFiveRxPzns%

Public LennartzPzn&
Public LennartzTxt$, LennartzDatei$

Public iind%


Private Const DefErrModul = "TAXIEREN.BAS"

Sub HoleTaxMusterZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleTaxMusterZeile")
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
Dim erg%

Select Case TmInhalt.flag
    Case MAG_HILFSTAXE
        erg% = CheckTaxierungHilfstaxe%
    Case MAG_SPEZIALITAET
        erg% = CheckTaxierungSpezialitaet%
    Case MAG_PREISEINGABE
        erg% = CheckTaxierungPreisEingabe%
    Case MAG_ARBEIT, MAG_GEFAESS
        erg% = CheckTaxierungArbEmb%
    Case MAG_SONSTIGES
        erg% = CheckTaxierungSonstiges
    Case MAG_ANTEILIG
        erg% = CheckTaxierungAnteilig%
    Case Else
        TmInhalt.flag = MAG_HILFSTAXE
End Select

If (erg% = False) Then
    With TaxierRec
        .pzn = TmInhalt.pzn
        .kurz = TmInhalt.kurz
        .menge = Space$(Len(.menge))
        .Meh = Space$(Len(.Meh))
        .kp = 0
        .GStufe = 0
        
        .ActMenge = TmInhalt.ActMenge
        .ActPreis = 0#
        
        .flag = TmInhalt.flag + MAG_NN
    End With
End If

Call DefErrPop
End Sub

Function CheckTaxierungHilfstaxe%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungHilfstaxe%")
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
Dim gramm#, gewicht#
Dim hStr$
Dim MatchHilfsTaxe As clsMatchHilfsTaxe
Dim MatchHilfsTaxeAdo As clsMatchHilfsTaxeADO

ret% = False

If (ArtikelDbOk) Then
    If (TmInhalt.pzn <> Space$(8)) And (TmInhalt.pzn <> "00000000") Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + TmInhalt.pzn
        FabsErrf = Hilfstaxe.OpenRecordset(HilfstaxeRec, SQLStr)
        If (FabsErrf% = 0) Then
            Call UmspeichernHilfsTaxe(True)
            TaxierRec.ActMenge = IIf(Val(TmInhalt.pzn) = 4443869, 1, TmInhalt.ActMenge * TmMengenFaktor#)
            Call PreisHilfsTaxe
            ret% = True
        End If
    Else
    '    hStr$ = UCase(Left$(TmInhalt.kurz, 10))
        Set MatchHilfsTaxeAdo = New clsMatchHilfsTaxeADO
        Call MatchHilfsTaxeAdo.InitMatch
        hStr$ = MatchHilfsTaxeAdo.MakeMatchCode(Left$(TmInhalt.kurz, 10))
        
        On Error Resume Next
        HilfstaxeRec.Close
        On Error GoTo 0

        SQLStr = "SELECT * FROM Artikel WHERE Komprimiert LIKE '" + Left(hStr, 2) + "%'"
        SQLStr = SQLStr + " ORDER BY SortierName"
        HilfstaxeRec.Open SQLStr$, Hilfstaxe.ActiveConn
        Do
            If (HilfstaxeRec.EOF) Then
                Exit Do
            End If
            
            Call Hilfstaxe.DB2DOS(HilfstaxeRec)
            If (hTaxe.kurz = TmInhalt.kurz) Then
                Call UmspeichernHilfsTaxe(True)
                TaxierRec.ActMenge = TmInhalt.ActMenge * TmMengenFaktor#
                Call PreisHilfsTaxe
                ret% = True
                Exit Do
            End If
            
            HilfstaxeRec.MoveNext
        Loop
    End If
Else
    If (TmInhalt.pzn <> Space$(7)) And (TmInhalt.pzn <> "0000000") Then
        FabsErrf% = hTaxe.IndexSearch(0, TmInhalt.pzn, FabsRecno&)
        If (FabsErrf% = 0) Then
            Call hTaxe.GetRecord(FabsRecno& + 1)
            Call UmspeichernHilfsTaxe(True)
            TaxierRec.ActMenge = TmInhalt.ActMenge * TmMengenFaktor#
            Call PreisHilfsTaxe
            ret% = True
        End If
    Else
    '    hStr$ = UCase(Left$(TmInhalt.kurz, 10))
        Set MatchHilfsTaxe = New clsMatchHilfsTaxe
        Call MatchHilfsTaxe.InitMatch
        hStr$ = MatchHilfsTaxe.MakeMatchCode(Left$(TmInhalt.kurz, 10))
        
        FabsErrf% = hTaxe.IndexGeneric(1, hStr$, FabsRecno&)
        If (FabsErrf% = 12) Then
            FabsErrf% = hTaxe.IndexNext(1, FabsRecno&, hStr$, FabsRecno&)
        ElseIf (FabsErrf% = 13) Or (FabsErrf% = 15) Then
            FabsErrf% = 0
        End If
        Do
            If (FabsErrf% <> 0) Then Exit Do
            
            Call hTaxe.GetRecord(FabsRecno& + 1)
            If (Left$(hTaxe.kurz, 2) <> Left$(TmInhalt.kurz, 2)) Then Exit Do
        
            If (hTaxe.kurz = TmInhalt.kurz) Then
                Call UmspeichernHilfsTaxe(True)
                TaxierRec.ActMenge = TmInhalt.ActMenge * TmMengenFaktor#
                Call PreisHilfsTaxe
                ret% = True
                Exit Do
            End If
        
            hStr$ = GetLastKey(10)
            FabsErrf% = hTaxe.IndexNext(1, FabsRecno&, hStr$, FabsRecno&)
        Loop
    End If
End If

CheckTaxierungHilfstaxe% = ret%

Call DefErrPop
End Function

Function CheckTaxierungSpezialitaet%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungSpezialitaet%")
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
Dim SQLStr$

ret% = False
      
If (ArtikelDbOk) Then
    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + TmInhalt.pzn
'                    SQLStr = SQLStr + " AND LagerKz<>0"
    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
Else
    FabsErrf% = ast.IndexSearch(0, TmInhalt.pzn, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
    End If
End If
If (FabsErrf% = 0) Then
'    ast.GetRecord (FabsRecno& + 1)
Else
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + TmInhalt.pzn
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If (TaxeRec.EOF = False) Then
        Call Taxe2ast(TmInhalt.pzn)
        FabsErrf% = 0
    End If
End If
    
If (FabsErrf% = 0) Then
    Call UmspeichernSpezialitaet
    TaxierRec.ActMenge = IIf(Val(TmInhalt.pzn) = 4443869, 1, TmInhalt.ActMenge * TmMengenFaktor#)

'    If (TaxierRec.ActMenge <= TaxierRec.Gstufe) Then
    If (ParenteralRezept >= 0) Or (TaxierRec.ActMenge <= TaxierRec.GStufe) Then
        Call PreisSpezialitaet(True)
        ret% = True
    End If
End If

CheckTaxierungSpezialitaet% = ret%

Call DefErrPop

End Function

Function CheckTaxierungPreisEingabe%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungPreisEingabe%")
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
      
Call UmspeichernPreisEingabe
TaxierRec.ActPreis = TmInhalt.ActPreis * TmMengenFaktor#

''neu in 5.0.5
'If (InStr(UCase(TmInhalt.kurz), "FIX-AUFSCHLAG") > 0) Then
'    With TaxierRec
'        .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
'        .ActPreis = TmInhalt.ActPreis
'    End With
'End If
Dim i%
Dim ParEnteralPeiseingaben$(2)
ParEnteralPeiseingaben$(0) = "Fix-Aufschlag"
ParEnteralPeiseingaben$(1) = "Fixzuschlag"
ParEnteralPeiseingaben$(2) = "BTM-Geb¸hr"
For i = 0 To 2
    If (InStr(UCase(TmInhalt.kurz), UCase(ParEnteralPeiseingaben$(i))) > 0) Then
        With TaxierRec
            .kurz = Left$(ParEnteralPeiseingaben$(i) + Space$(Len(.kurz)), Len(.kurz))
            .ActPreis = TmInhalt.ActPreis
            Exit For
        End With
    End If
Next i


'Call PreisSpezialitaet

CheckTaxierungPreisEingabe% = True

Call DefErrPop
End Function

Function CheckTaxierungArbEmb%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungArbEmb%")
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
Dim i%, ind%, NeuActMenge%, ret%
Dim h$, SollTyp$

ret% = False
   
If (ParenteralRezept >= 0) And (TmInhalt.flag = MAG_ARBEIT) Then
    For i% = 0 To UBound(ParenteralPzn)
        h$ = ParenteralPzn(i)
        If (h = TmInhalt.pzn) Then
            With TaxierRec
                .pzn = ParenteralPzn(i)
                .kurz = ParenteralTxt(i)
                
                .menge = Space$(Len(.menge))
                .Meh = Space$(Len(.Meh))
                
                If (Parenteral_AOK_LosGebiet) Then
                    .kp = 0
                Else
                    .kp = ParenteralPreis(i)
                End If
                
                .GStufe = TmInhalt.ActMenge
                
                .ActMenge = .GStufe
                .ActPreis = .kp
                
                .flag = MAG_ARBEIT
            End With
            CheckTaxierungArbEmb% = True
            Call DefErrPop: Exit Function
        End If
    Next i
End If

If (ParenteralRezept >= 0) Then
    NeuActMenge% = TmInhalt.ActMenge
Else
    NeuActMenge% = TmHeader.ActMenge * TmMengenFaktor#
End If
h$ = iTrim$(TmInhalt.kurz)
TmInhalt.kurz = h$

If (RezRbhLauer) Then
    Dim OrgMenge As String
    OrgMenge = ""
    
     If (TmInhalt.flag = MAG_GEFAESS) Then
        h = Trim(TmInhalt.kurz)
        If (Val(TmInhalt.pzn) > 0) Then
            Dim pzn$
            pzn = TmInhalt.pzn
            
            If (ArtikelDbOk) Then
                SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
'                    SQLStr = SQLStr + " AND LagerKz<>0"
                FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
            Else
                FabsErrf% = ast.IndexSearch(0, pzn, FabsRecno&)
                If (FabsErrf% = 0) Then
                    ast.GetRecord (FabsRecno& + 1)
                End If
            End If
            If (FabsErrf% = 0) Then
'                            ast.GetRecord (FabsRecno& + 1)
            Else
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    Call Taxe2ast(pzn$)
                    FabsErrf% = 0
                End If
            End If
                
            If (FabsErrf% = 0) Then
                Call UmspeichernSpezialitaet(MAG_GEFAESS)
                TaxierRec.ActPreis = TaxierRec.kp
                TaxierRec.ActMenge = TaxierRec.GStufe
                If (ParenteralRezept = 24) Or (ParenteralRezept = 26) Or (ParenteralRezept = 28) Or (ParenteralRezept = 30) Then
                    TaxierRec.ActPreis = TaxierRec.ActPreis * 1.9
                ElseIf (ParenteralRezept = 25) Or (ParenteralRezept = 27) Or (ParenteralRezept = 29) Then
                    TaxierRec.ActPreis = TaxierRec.ActPreis * 2
                Else
                    TaxierRec.ActPreis = TaxierRec.ActPreis * MalFaktor#
                End If
                
                CheckTaxierungArbEmb = True
                Call DefErrPop: Exit Function
            End If


            SQLStr = "SELECT * FROM Artikel WHERE PZN = " + TmInhalt.pzn
            RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
            If Not (RezRbhLauerRec.EOF) Then
                h = CheckNullStr(RezRbhLauerRec!Name)
                OrgMenge = Trim(CheckNullStr(RezRbhLauerRec!menge))
            End If
            RezRbhLauerRec.Close
        End If
        
        SQLStr = "SELECT TOP 1 Id,PZN,Name,FloatMenge AS Menge,Einheit,Kp FROM Artikel AS A"
        SQLStr = SQLStr + " WHERE (SUBSTRING(Name,1," + CStr(Len(h)) + ")='" + h + "') AND (FloatMenge>=" + uFormat(NeuActMenge, "0.00") + ")"
        SQLStr = SQLStr + " AND (Emballage<>0)"
        SQLStr = SQLStr + " ORDER BY Menge"
        RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
        If RezRbhLauerRec.EOF And (OrgMenge > "") Then
            'manchmal steht Menge im Namen -> k¸rzen
            i = InStr(h, " " & OrgMenge)
            If i > 0 Then
                h = Trim(Left(h, i - 1))
                SQLStr = "SELECT TOP 1 Id,PZN,Name,FloatMenge AS Menge,Einheit,Kp FROM Artikel AS A"
                SQLStr = SQLStr + " WHERE (SUBSTRING(Name,1," + CStr(Len(h)) + ")='" + h + "') AND (FloatMenge>=" + uFormat(NeuActMenge, "0.00") + ")"
                SQLStr = SQLStr + " AND (Emballage<>0)"
                SQLStr = SQLStr + " ORDER BY Menge"
                RezRbhLauerRec.Close
                RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
            End If
        End If
     Else
        SQLStr = "SELECT TOP 1 * FROM ArbeitsPreise AS A"
        SQLStr = SQLStr + " WHERE (Arbeit='" + TmInhalt.kurz + "') AND (Menge>=" + uFormat(NeuActMenge, "0.00") + ")"
        SQLStr = SQLStr + " ORDER BY Menge"
        RezRbhLauerRec.Open SQLStr, taxeAdoDB.ActiveConn
    End If
    
    Do
        If (RezRbhLauerRec.EOF) Then
            Exit Do
        End If
        
        Call UmspeichernArbEmb(TmInhalt.flag)
        TaxierRec.ActPreis = TaxierRec.kp
        
        If (TmInhalt.flag = MAG_GEFAESS) Then
            TaxierRec.ActPreis = TaxierRec.ActPreis * MalFaktor#
        ElseIf (Left$(TmInhalt.kurz, 5) = "UNVER") Then
            NeuMalFaktor# = 2#
        Else
            NeuMalFaktor = 1.9
        End If
    
        TaxierRec.ActMenge = TaxierRec.GStufe
                        
        ret% = True
            
        Exit Do
        
        RezRbhLauerRec.MoveNext
    Loop
    RezRbhLauerRec.Close
Else
    If (TmInhalt.flag = MAG_GEFAESS) Then
        ind% = AnfEmb%
        SollTyp$ = "E"
    Else
        ind% = 1
        SollTyp$ = "A"
    End If
    
    Seek #ARBEMB%, (128& * ind%) + 1
            
    h$ = iTrim$(TmInhalt.kurz)
    TmInhalt.kurz = h$
    
    Do
        Get #ARBEMB%, , ArbEmbRec
        If (EOF(ARBEMB%)) Then Exit Do
        If (ArbEmbRec.Typ <> SollTyp$) Then Exit Do
    
        Call UmspeichernArbEmb(TmInhalt.flag)
    
        If (ParenteralRezept >= 0) Then
            NeuActMenge% = TmInhalt.ActMenge
        Else
            NeuActMenge% = TmHeader.ActMenge * TmMengenFaktor#
        End If
        If (TaxierRec.GStufe >= NeuActMenge%) And (TaxierRec.kurz = TmInhalt.kurz) Then
            TaxierRec.ActPreis = TaxierRec.kp
            
            If (TmInhalt.flag = MAG_GEFAESS) Then
                TaxierRec.ActPreis = TaxierRec.ActPreis * MalFaktor#
            ElseIf (Left$(TmInhalt.kurz, 5) = "UNVER") Then
                NeuMalFaktor# = 2#
            Else
                NeuMalFaktor = 1.9
            End If
        
            TaxierRec.ActMenge = TaxierRec.GStufe
            
            ret% = True
            
            Exit Do
        End If
    Loop
End If
        
CheckTaxierungArbEmb% = ret%

Call DefErrPop
End Function

Function CheckTaxierungSonstiges%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungSonstiges%")
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
Dim ind%, NeuActMenge%, ret%
Dim h$, SollTyp$

ret% = False
   
ind% = 1
SollTyp$ = "A"

Seek #ARBEMB%, (128& * ind%) + 1
        
h$ = iTrim$(TmInhalt.kurz)
TmInhalt.kurz = h$

'h$ = iTrim$(TmInhalt.kurz)
'h$ = Left$(h$ + Space$(28), 28)

Do
    Get #ARBEMB%, , ArbEmbRec
    If (EOF(ARBEMB%)) Then Exit Do
    If (ArbEmbRec.Typ <> SollTyp$) Then Exit Do

    Call UmspeichernArbEmb(TmInhalt.flag)

    If (TaxierRec.GStufe = 0) And (TaxierRec.kurz = TmInhalt.kurz) Then
        TaxierRec.ActMenge = TmInhalt.ActMenge
        Call PreisSonstiges
        ret% = True
        Exit Do
    End If
Loop
        
CheckTaxierungSonstiges% = ret%

Call DefErrPop
End Function

Function CheckTaxierungAnteilig%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckTaxierungAnteilig%")
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
Dim SQLStr$

ret% = False
      
If (ArtikelDbOk) Then
    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + TmInhalt.pzn
'                    SQLStr = SQLStr + " AND LagerKz<>0"
    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
Else
    FabsErrf% = ast.IndexSearch(0, TmInhalt.pzn, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
    End If
End If
If (FabsErrf% = 0) Then
'    ast.GetRecord (FabsRecno& + 1)
Else
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + TmInhalt.pzn
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If (TaxeRec.EOF = False) Then
        Call Taxe2ast(TmInhalt.pzn)
        FabsErrf% = 0
    End If
End If
    
If (FabsErrf% = 0) Then
    Call UmspeichernSpezialitaet(MAG_ANTEILIG)
    TaxierRec.ActMenge = TmInhalt.ActMenge * TmMengenFaktor#
    If (TaxierRec.ActMenge <= TaxierRec.GStufe) Then
        Call PreisAnteilig
        ret% = True
    End If
End If

CheckTaxierungAnteilig% = ret%

Call DefErrPop
End Function

Sub UmspeichernHilfsTaxe(Optional CheckPreis0% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UmspeichernHilfsTaxe")
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

With TaxierRec
    .pzn = hTaxe.pzn
    .kurz = hTaxe.kurz
    .menge = hTaxe.meng
    .Meh = hTaxe.Meh
    .kp = hTaxe.kp
    .GStufe = Val(hTaxe.GStufe)
    
    .ActMenge = 0#
    .ActPreis = 0#
    
    .flag = MAG_HILFSTAXE

    If (CheckPreis0% And (.kp = 0)) Then
        Call MessageBox("Achtung: kein Preis vorhanden !", vbInformation, .pzn + "  " + .kurz)
    End If
End With


Call DefErrPop
End Sub

Sub UmspeichernSpezialitaet(Optional tTyp As Byte = MAG_SPEZIALITAET)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UmspeichernSpezialitaet")
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
Dim sMenge$, sEinheit$
'Dim tRec2 As Recordset
Dim tRec2 As New ADODB.Recordset

Dim dAstAep#, dTaxeEk#
Dim bRezeptpflichtig As Boolean


With TaxierRec
    .pzn = ast.pzn
    .kurz = ast.kurz
    .menge = ast.meng
    .Meh = ast.Meh
    .kp = ast.aep
    
    dAstAep = ast.aep
    bRezeptpflichtig = ((Left(ast.rez, 1) = "+"))
    dTaxeEk = 0
    
'    .Gstufe = Val(ast.meng)
    sEinheit = ast.Meh
    sMenge$ = ast.meng
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ast.pzn
'    Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    tRec2.Close
    Err.Clear
    On Error GoTo DefErr

    tRec2.Open SQLStr, taxeAdoDB.ActiveConn
    If (tRec2.EOF = False) Then
        sMenge$ = tRec2!menge
        dTaxeEk = CheckNullLong(tRec2!EK) / 100#
                        
        Dim iVal%
        iVal% = CheckNullInt(tRec2!Abgabebest)
        bRezeptpflichtig = (iVal% = 1) Or (iVal% = 5) Or (iVal% = 6)
        
        If (tTyp = MAG_ANTEILIG) Then
            If (dTaxeEk > 0) Then
                .kp = dTaxeEk
            End If
        End If
    End If

    If (ParenteralRezept >= 0) And (ParenteralRezept <= 15) Then
    Else
        Dim sSQL$, pzn&
        Dim dDichte#
        
        dDichte = 0
        pzn = 0
    
#If (WINREZX = 1) Then
#Else
        sSQL = "Select GRP_HA.*, HA.Dichte FROM GRP_HA LEFT JOIN HA ON GRP_HA.PZN_2=HA.PZN WHERE GRP_HA.PZN_1=" + .pzn
        On Error Resume Next
        ABDA_Komplett_Rec.Close
        Err.Clear
        On Error GoTo DefErr
        ABDA_Komplett_Rec.Open sSQL, ABDA_Komplett_Conn
        If (ABDA_Komplett_Rec.EOF = False) Then
            pzn = CheckNullLong(ABDA_Komplett_Rec!PZN_2)
            dDichte = xVal(CheckNullStr(ABDA_Komplett_Rec!Dichte))
            If (dDichte = 0) Then
                dDichte = 1
            Else
                dDichte = dDichte / 10000
            End If
        End If
        ABDA_Komplett_Rec.Close
#End If
        If (pzn > 0) Then
            Dim i%
            Dim GStufe#(1), Meh$(1), h$

'            .Dichte = dDichte

            SollPzn = .pzn
            SollTaxierTyp = tTyp
'            frmMarktPzn.Show 1

            h = "FAM: " + vbCrLf + .pzn + " " + .kurz + " " + .menge + " " + .Meh + vbCrLf
            h = h + "EK: " + Format(.kp, "0.00") + vbCrLf

            Meh(0) = sEinheit
            GStufe(0) = GPMenge(Meh(0), sMenge$)
            .Meh = Meh(0)

            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(pzn)
            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
            If (FabsErrf% = 0) Then
            '    ast.GetRecord (FabsRecno& + 1)
            Else
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(pzn)
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    Call Taxe2ast(PznString(pzn))
                    FabsErrf% = 0
                End If
            End If
            If (FabsErrf% = 0) Then
'                AuswahlHaPzn = AdvReader2(AdvReader2.GetOrdinal("Pzn"))

                h = h & vbCrLf + "HA:" + vbCrLf
                h = h & " " + ast.pzn
                h = h & " " + ast.pzn
                h = h & " " + ast.meng
                h = h & " " + ast.Meh + vbCrLf
                h = h & "EK: " + Format(ast.aep, "0.00") + vbCrLf ' CheckNullDouble(RezRbhLauerRec!Preis)

                Meh(1) = ast.Meh
                GStufe(1) = GPMenge(Meh(1), ast.MengNeu)

                For i = 0 To 1
                    If (Meh(i) = "G") And (Meh((i + 1) Mod 2) = "MG") Then
                        GStufe(i) = GStufe(i) * 1000#
                        Meh(i) = "MG"
                        Exit For
                    End If
                Next

                If (tTyp = MAG_GEFAESS) Then
                    .Meh = "ST"
                    sMenge = "1"
                    GStufe(0) = 1
                    Meh(0) = "ST"
                    GStufe(1) = 1
                    Meh(1) = "ST"
                End If

                h = h & vbCrLf + "Mengen:" + vbCrLf
                For i = 0 To 1
                    h = h & IIf(i = 0, "FAM", "HA") + ":  " + CStr(GStufe(i)) + " " + Meh(i) + vbCrLf
                Next

                If (Meh(0) <> Meh(1)) Then
                    '                MsgBox("<" + Meh(0) + ">  <" + Meh(1) + ">")
                    For i = 0 To 1
                        If (Meh(i) = "ML") And (Meh((i + 1) Mod 2) = "G") Then
                            GStufe(i) = GStufe(i) * dDichte
                            Meh(i) = "G"
                            '                                        h &= vbCrLf + vbCrLf + "Dichte: " + dDichte.ToString + "  GStufe: " + GStufe(i).ToString + "   Meh: " + Meh(i)
                            Exit For
                        End If
                    Next
                    h = h & vbCrLf + "Dichte: " + CStr(dDichte) + vbCrLf
                    For i = 0 To 1
                        h = h & IIf(i = 0, "FAM", "HA") + ":  " + CStr(GStufe(i)) + " " + Meh(i) + vbCrLf
                    Next
                End If

                'Dim Faktor As Double = (.Gstufe / AdvReader2(AdvReader2.GetOrdinal("Menge")))
                Dim Faktor As Double
                Faktor = GStufe(0) / GStufe(1)
                .kp = ast.aep * Faktor
                'MsgBox(AdvReader2(AdvReader2.GetOrdinal("ApothekenEk1")).ToString + " " + Faktor.ToString + " " + (AdvReader2(AdvReader2.GetOrdinal("ApothekenEk1")) * Faktor).ToString + " " + .kp.ToString)

                h = h & vbCrLf + "Mengen-Faktor:  " + Format(Faktor, "0.00")
                h = h & vbCrLf + "FAM-EK neu:  " + Format(.kp, "0.00")
            End If
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr

            h = h & vbCrLf + vbCrLf

'            MsgBox (h)
            
'            .AI = 1
        End If
        
        If (Year(Now) >= 2024) Then
            .kp = IIf(bRezeptpflichtig, dTaxeEk, dAstAep)
        End If
    
    End If
        
    .GStufe = GPMenge(sEinheit$, sMenge$)
    .Meh = sEinheit
    
    .ActMenge = 0#
    .ActPreis = 0#
    
    .flag = tTyp ' MAG_SPEZIALITAET
End With

Call DefErrPop
End Sub

Sub UmspeichernPreisEingabe()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UmspeichernPreisEingabe")
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

With TaxierRec
    .pzn = Space$(Len(.pzn))
    .kurz = Left$("PREIS-EINGABE" + Space$(Len(.kurz)), Len(.kurz))
    .menge = Space$(Len(.menge))
    .Meh = Space$(Len(.Meh))
    .kp = 0
    .GStufe = 0
    
    .ActMenge = 0#
    .ActPreis = 0#
    
    .flag = MAG_PREISEINGABE
End With

Call DefErrPop
End Sub

Sub UmspeichernArbEmb(SollFlag As Byte)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UmspeichernArbEmb")
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
Dim a!
Dim b#

With TaxierRec
    
    .pzn = Space$(Len(.pzn))
    
    If (RezRbhLauer) Then
        .menge = Space$(Len(.menge))
        .Meh = Space$(Len(.Meh))  'ArbEmbRec.meh
        .GStufe = CheckNullDouble(RezRbhLauerRec!menge)
        If (SollFlag = MAG_GEFAESS) Then
            .kurz = iTrim(CheckNullStr(RezRbhLauerRec!Name))
            .kp = CheckNullDouble(RezRbhLauerRec!kp)
            .pzn = CheckNullLong(RezRbhLauerRec!pzn)
        Else
            .kurz = iTrim(CheckNullStr(RezRbhLauerRec!Arbeit))
            .kp = CheckNullDouble(RezRbhLauerRec!Preis)
        End If
    Else
        .kurz = iTrim(ArbEmbRec.kurz)
        
        .menge = Space$(Len(.menge))
        .Meh = ArbEmbRec.Meh
        
        b# = ArbEmbRec.kp
        Call DxToIEEEd(b#)
        .kp = b# / 100#
        
        a! = ArbEmbRec.menge
        Call DxToIEEEs(a!)
        .GStufe = a!
    End If
    
    .ActMenge = 0#
    .ActPreis = 0#
    
    .flag = SollFlag
End With

Call DefErrPop
End Sub

Sub PreisHilfsTaxe()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisHilfsTaxe")
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
Dim gramm#, gewicht#

With TaxierRec
    gramm# = .ActMenge
    gewicht# = .GStufe
    If (gewicht# = 0#) Then
        .ActPreis = 0#
    Else
        .ActPreis = .kp * gramm# / gewicht#
        If (.ActPreis < 0.01) Then .ActPreis = 0.01
        .ActPreis = .ActPreis * MalFaktor#
    End If
End With

Call DefErrPop
End Sub

'Sub PreisSpezialitaet(Optional pEk# = 0, Optional pPreisProEinheit# = 0, Optional pAnzEinheiten# = 0)
Sub PreisSpezialitaet(Optional TaxMusterAktiv% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisSpezialitaet")
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
Dim gramm#, gewicht#, CentProEinheit#, PreisFaktor#
Dim h$, MgPreiseMDB$
Dim Anlage3Ok%
Dim Anlage3DB As Database
Dim Anlage3Rec As Recordset

With TaxierRec
    gramm# = .ActMenge
    gewicht# = .GStufe
    If (ParenteralRezept >= 0) Then
    ElseIf (gramm# > gewicht#) Then
        .ActMenge = .GStufe
    End If
    If (gramm# = 0) Then
        .ActPreis = 0
    Else
'        If (pEk > 0) Then
'            Load frmEdit
'            With frmEdit.lstEdit
'                .Left = 0
'                .Top = 0
'                .Width = frmTaxieren.flxTaxieren.Width
'                .Height = frmTaxieren.flxTaxieren.RowHeight(0) * 5
'
'                .Clear
'                .AddItem "0  nicht patentgesch¸tzt (Abschlag von 10% des EK)"
'                .AddItem "1  nicht patentgesch¸tzt, kein anderes FAM abgebbar (Abschlag von 1% des EK)"
'                .AddItem "2  andere nicht patent-gesch¸tzte Arzneimittel"
'                .AddItem "3  patentgesch¸tzt"
'
'                .Visible = True
'            End With
'            With frmEdit
'                .Left = frmTaxieren.flxTaxieren.Left '+ 45 '+ flxTaxieren.ColPos(col%) + 45
'                .Left = .Left + frmTaxieren.Left + wpara.FrmBorderHeight
'                .Top = frmTaxieren.flxTaxieren.Top + frmTaxieren.flxTaxieren.RowPos(frmTaxieren.flxTaxieren.row)
'                .Top = .Top + frmTaxieren.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
'                .Width = frmTaxieren.flxTaxieren.ColWidth(0)
'                .Width = frmTaxieren.flxTaxieren.Width '- flxTaxieren.ColPos(col%) + 45
'                .Height = .lstEdit.Height
'            End With
'
'            frmEdit.Show 1
'
'            If (EditErg%) Then
'                EditErg = Val(Left(EditTxt, 1))
''                MsgBox (Str(EditErg) + Str(.ActMenge) + Str(.Gstufe))
'
'                .Kp = pPreisProEinheit * (pAnzEinheiten * .ActMenge / .Gstufe)
'                If (EditErg = 0) Then
'                    .Kp = .Kp * 0.9
'                ElseIf (EditErg = 1) Then
'                    .Kp = .Kp * 0.99
'                End If
'            End If
'        End If
        .ActPreis = .kp
        If (ParenteralRezept = 28) Or (ParenteralRezept = 29) Or (ParenteralRezept = 30) Then
            .ActPreis = (.kp * .ActMenge) / .GStufe
            If (ParenteralRezept = 28) Then
                If (InStr(.Meh, "G") > 0) Then
                    Do
                        h$ = InputBox("Dichte: ", "Cannabisextrakt", "1")
                        h$ = UCase(Trim(h$))
                        If (h$ = "") Then
                            Exit Do
                        ElseIf (xVal(h$) > 0) Then
                            CannabisExtraktDichte = xVal(h$)
                            Exit Do
                        End If
                    Loop
                    If (CannabisExtraktDichte = 0) Then
                        CannabisExtraktDichte = 1
                    End If
                    h = Format(.ActMenge / CannabisExtraktDichte, "0.00")
                    Dim CannabisExtraktVerordnetML#
                    Do
                        h$ = InputBox("Verordnete Menge in ML: ", "Cannabisextrakt", h)
                        h$ = UCase(Trim(h$))
                        If (h$ = "") Then
                            CannabisExtraktVerordnetML = .ActMenge / CannabisExtraktDichte
                            Exit Do
                        ElseIf (xVal(h$) > 0) Then
                            CannabisExtraktVerordnetML = xVal(h$)
                            Exit Do
                        End If
                    Loop
                    .ActMenge = CannabisExtraktVerordnetML * CannabisExtraktDichte
                    .ActPreis = (.kp * .ActMenge) / .GStufe
                End If
            End If
        ElseIf (ParenteralRezept > 15) And (ParenteralRezept < 24) Then
            .ActPreis = 0
        ElseIf (ParenteralRezept = 24) Or (ParenteralRezept = 25) Or (ParenteralRezept = 26) Or (ParenteralRezept = 27) Then
            .ActPreis = 0
        ElseIf (ParenteralRezept >= 0) And (ParenteralRezept <= 15) Then
            If (Parenteral_AOK_LosGebiet) Then
                MgPreiseMDB = "MgPreis1.mdb"
            ElseIf (Parenteral_AOK_NordOst) Then
                MgPreiseMDB = "MgPreis2.mdb"
            Else
                MgPreiseMDB = "MgPreise.mdb"
            End If
            If (Dir(MgPreiseMDB) <> "") Then
                Set Anlage3DB = OpenDatabase(MgPreiseMDB, False, True)
                
                SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + .pzn
                Set Anlage3Rec = Anlage3DB.OpenRecordset(SQLStr$)
                If Not (Anlage3Rec.EOF) Then
                
                    CentProEinheit = Anlage3Rec!StoffpreisNonAI * Anlage3Rec!FaktorNonAI ' Anlage3Rec!EinheitspreisNonAI
'                    PreisFaktor = Anlage3Rec!FaktorNonAI
                    If (TaxMusterAktiv) Then
                        If (ParEnteralAI) Then
                            CentProEinheit = Anlage3Rec!StoffpreisAI * Anlage3Rec!FaktorAI ' Anlage3Rec!EinheitspreisAI
'                            PreisFaktor = Anlage3Rec!FaktorAI
                        End If
                    Else
                        If (Anlage3Rec!EinheitspreisAI <> Anlage3Rec!EinheitspreisNonAI) Then
                            If (MessageBox("AutIdem zu diesem Artikel zul‰ssig?", vbYesNo Or vbDefaultButton1) = vbYes) Then
                                CentProEinheit = Anlage3Rec!StoffpreisAI * Anlage3Rec!FaktorAI ' Anlage3Rec!EinheitspreisAI
'                                PreisFaktor = Anlage3Rec!FaktorAI
                                ParEnteralAI = True
                            End If
                        End If
                        
                        ParEnteralAnzEinheiten = 0
                        Do
                            h$ = MyInputBox("Wirkstoffmenge: ", "Parenterale Taxierung", "")
                            h$ = UCase(Trim(h$))
                            If (h$ = "") Then
                                Exit Do
                            ElseIf (xVal(h$) > 0) Then
                                ParEnteralAnzEinheiten = xVal(h$)
                                Exit Do
                            End If
                        Loop
                    End If
                    
                    .ActPreis = ParEnteralAnzEinheiten * CentProEinheit / 100
'                    .ActPreis = ParEnteralAnzEinheiten * CentProEinheit * PreisFaktor / 100
                ElseIf (TaxMusterAktiv = 0) Then
                    ParEnteralPrim‰rPackmittel = (MessageBox("Handelt es sich um ein Prim‰rpackmittel?", vbYesNo Or vbDefaultButton2) = vbYes)
                    If Not (ParEnteralPrim‰rPackmittel) And (Parenteral_AOK_LosGebiet Or Parenteral_AOK_NordOst) Then
                        Call MessageBox("Achtung: Artikel nicht zul‰ssig", vbCritical)
                        .ActMenge = 0
                        .ActPreis = 0
                    End If
                End If
                
                Anlage3DB.Close
            End If
        End If
'        .ActPreis = .Kp
        If (ParEnteralPrim‰rPackmittel) Then
            If (Parenteral_AOK_LosGebiet) Then
                .ActPreis = 0
            ElseIf (.GStufe > 1) Then
                .ActPreis = (.kp * .ActMenge) / .GStufe
            End If
        End If
    End If
    .ActPreis = .ActPreis * MalFaktor#
End With

Call DefErrPop
End Sub

Sub PreisSonstiges()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisSonstiges")
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
Dim gramm#, gewicht#

With TaxierRec
    If (.ActMenge > 1000) Then
        .ActPreis = 0#
    Else
        .ActPreis = .kp * .ActMenge
    End If
End With

Call DefErrPop
End Sub

Sub PreisAnteilig()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisAnteilig")
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
Dim gramm#, gewicht#

With TaxierRec
    gramm# = .ActMenge
    gewicht# = .GStufe
    If (gramm# > gewicht#) Then
        gramm# = gewicht#
        .ActMenge = gewicht#
    End If
    If (gramm# = 0) Then
        .ActPreis = 0
    Else
        .ActPreis = .kp * gramm# / gewicht#
        If (.ActPreis < 0.01) Then .ActPreis = 0.01
        .ActPreis = .ActPreis * MalFaktor#
    End If
End With

Call DefErrPop
End Sub

Sub OpenArbEmb()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenArbEmb")
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
Dim iAnz%
Dim h$
      
      
' SQL-Statement zum Pr¸fen auf Vorhandensein des Feldes
Dim lRecs&, sField$
sField = "Emballage"
SQLStr = "SELECT * FROM Information_Schema.Columns WHERE Table_Name LIKE 'ArbeitsPreise'" 'AND Column_Name LIKE '" & sField & "'"
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
RezRbhLauer = (TaxeRec.EOF = False)
TaxeRec.Close
'RezRbhLauer = 0
      
If (RezRbhLauer) Then
    Call RezRbhKonvert
Else
    If (Dir$("rezrbh.dat") = "") Then Call DefErrPop: Exit Sub
    
    ARBEMB% = FileOpen("rezrbh.dat", "R", "B")
    If (ARBEMB% <= 0) Then
        Call DefErrPop: Exit Sub
    End If
    
    Seek #ARBEMB%, 1
    h$ = String(2, 0)
    Get #ARBEMB%, , h$
    iAnz% = CVI(h$)
    
'        Dim F_OUT%, a!
'        F_OUT = FreeFile
'        Open "f_out.txt" For Output As #F_OUT
'        Seek #ARBEMB%, (128&) + 1
'        Do
'            Get #ARBEMB%, , ArbEmbRec
'            If (EOF(ARBEMB%)) Then Exit Do
'            If (ArbEmbRec.Typ = "E") Then
'                a! = ArbEmbRec.menge
'                Call DxToIEEEs(a!)
'
'                Print #F_OUT, ArbEmbRec.kurz + " " + CStr(a)
'            End If
'        Loop
'        Close #F_OUT
        
    
    
    
    AnfArb% = 1
    Seek #ARBEMB%, (128& * AnfArb%) + 1
    Do
        Get #ARBEMB%, , ArbEmbRec
        If (Left$(ArbEmbRec.kurz, 1) <> Chr$(0)) Then Exit Do
        AnfArb% = AnfArb% + 1
    Loop
    
    AnfEmb% = iAnz% / 2
    Seek #ARBEMB%, (128& * AnfEmb%) + 1
    Get #ARBEMB%, , ArbEmbRec
    If (ArbEmbRec.Typ = "E") Then
        Do
            AnfEmb% = AnfEmb% - 1
            Seek #ARBEMB%, (128& * AnfEmb%) + 1
            Get #ARBEMB%, , ArbEmbRec
            If (ArbEmbRec.Typ = "A") Then
                AnfEmb% = AnfEmb% + 1
                Exit Do
            End If
        Loop
    Else
        Do
            AnfEmb% = AnfEmb% + 1
            Seek #ARBEMB%, (128& * AnfEmb%) + 1
            Get #ARBEMB%, , ArbEmbRec
            If (ArbEmbRec.Typ = "E") Then Exit Do
        Loop
    End If
End If

Call DefErrPop
End Sub

Sub RezRbhKonvert()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezRbhKonvert")
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
Dim LauerEmballagePzn&()
Dim LauerEmballageOpName$()
Dim LauerEmballageOpMenge#()
Dim i%, iAnz%, gef%
Dim HSatz&, lRecs&
Dim a!, kp#
Dim h$, sKurz$, Komprimiert$, SortierName$
Dim MatchArt As clsMatchArtikelADO

If (Dir("rezrbh.dat") = "") Then
    Call DefErrPop: Exit Sub
End If

ARBEMB% = FileOpen("rezrbh.dat", "R", "B")

iAnz = 72
ReDim Preserve LauerEmballagePzn(iAnz)
ReDim Preserve LauerEmballageOpName(iAnz)
ReDim Preserve LauerEmballageOpMenge(iAnz)

i = 0

LauerEmballagePzn(i) = 2182293: LauerEmballageOpName(i) = "APONORM DREHDOSIERKRUKEN": LauerEmballageOpMenge(i) = 20: i = i + 1
LauerEmballagePzn(i) = 2182301: LauerEmballageOpName(i) = "APONORM DREHDOSIERKRUKEN": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 2182318: LauerEmballageOpName(i) = "APONORM DREHDOSIERKRUKEN": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2182324: LauerEmballageOpName(i) = "APONORM DREHDOSIERKRUKEN": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 2182330: LauerEmballageOpName(i) = "APONORM DREHDOSIERKRUKEN": LauerEmballageOpMenge(i) = 200: i = i + 1
LauerEmballagePzn(i) = 2598728: LauerEmballageOpName(i) = "AUGENTROPFGLAS STERIL": LauerEmballageOpMenge(i) = 10: i = i + 1
LauerEmballagePzn(i) = 6347377: LauerEmballageOpName(i) = "EINZELDOSISBEHAELTER": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599774: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 0": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599768: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 00": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599780: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 1": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599805: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 2": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2598881: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 3": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2598898: LauerEmballageOpName(i) = "GELATINEKAPSEL GR. 4": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599024: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2599030: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 2599053: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 150: i = i + 1
LauerEmballagePzn(i) = 2599076: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 200: i = i + 1
LauerEmballagePzn(i) = 2599082: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 250: i = i + 1
LauerEmballagePzn(i) = 2599099: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 300: i = i + 1
LauerEmballagePzn(i) = 2599107: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 500: i = i + 1
LauerEmballagePzn(i) = 2599113: LauerEmballageOpName(i) = "GEWINDEFLASCHE GL 28": LauerEmballageOpMenge(i) = 1000: i = i + 1
LauerEmballagePzn(i) = 2598906: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 10: i = i + 1
LauerEmballagePzn(i) = 2598912: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 20: i = i + 1
LauerEmballagePzn(i) = 2598929: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 2598935: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2598941: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 75: i = i + 1
LauerEmballagePzn(i) = 2598958: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 2598964: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 150: i = i + 1
LauerEmballagePzn(i) = 2598970: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 200: i = i + 1
LauerEmballagePzn(i) = 2598987: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 250: i = i + 1
LauerEmballagePzn(i) = 2598993: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 300: i = i + 1
LauerEmballagePzn(i) = 2599001: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 500: i = i + 1
LauerEmballagePzn(i) = 2599018: LauerEmballageOpName(i) = "KRUKE MIT DECKEL WEISS KST": LauerEmballageOpMenge(i) = 1000: i = i + 1
LauerEmballagePzn(i) = 6179804: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 20: i = i + 1
LauerEmballagePzn(i) = 7332337: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 6179810: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 6179827: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 7332343: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 200: i = i + 1
LauerEmballagePzn(i) = 246712: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 300: i = i + 1
LauerEmballagePzn(i) = 246712: LauerEmballageOpName(i) = "KRUKE UNGUATOR": LauerEmballageOpMenge(i) = 500: i = i + 1
LauerEmballagePzn(i) = 2599917: LauerEmballageOpName(i) = "NASENSPRAY-FLASCHE": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599219: LauerEmballageOpName(i) = "PIPETTENGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 10: i = i + 1
LauerEmballagePzn(i) = 2599231: LauerEmballageOpName(i) = "PIPETTENGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 20: i = i + 1
LauerEmballagePzn(i) = 2599248: LauerEmballageOpName(i) = "PIPETTENGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 2599254: LauerEmballageOpName(i) = "PIPETTENGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2599260: LauerEmballageOpName(i) = "PIPETTENGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 2182347: LauerEmballageOpName(i) = "REZEPTURDOSE F TOPITEC": LauerEmballageOpMenge(i) = 300: i = i + 1
LauerEmballagePzn(i) = 2182382: LauerEmballageOpName(i) = "REZEPTURDOSE F TOPITEC": LauerEmballageOpMenge(i) = 500: i = i + 1
LauerEmballagePzn(i) = 2599432: LauerEmballageOpName(i) = "TROPFGLAS BRAUN RUND KOMPL": LauerEmballageOpMenge(i) = 10: i = i + 1
LauerEmballagePzn(i) = 2599455: LauerEmballageOpName(i) = "TROPFGLAS BRAUN RUND KOMPL": LauerEmballageOpMenge(i) = 20: i = i + 1
LauerEmballagePzn(i) = 2599461: LauerEmballageOpName(i) = "TROPFGLAS BRAUN RUND KOMPL": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 2599478: LauerEmballageOpName(i) = "TROPFGLAS BRAUN RUND KOMPL": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2599490: LauerEmballageOpName(i) = "TROPFGLAS BRAUN RUND KOMPL": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 7536666: LauerEmballageOpName(i) = "TUBE": LauerEmballageOpMenge(i) = 15: i = i + 1
LauerEmballagePzn(i) = 2599811: LauerEmballageOpName(i) = "TUBE": LauerEmballageOpMenge(i) = 25: i = i + 1
LauerEmballagePzn(i) = 2599828: LauerEmballageOpName(i) = "TUBE": LauerEmballageOpMenge(i) = 35: i = i + 1
LauerEmballagePzn(i) = 2599834: LauerEmballageOpName(i) = "TUBE": LauerEmballageOpMenge(i) = 60: i = i + 1
LauerEmballagePzn(i) = 2599840: LauerEmballageOpName(i) = "TUBE": LauerEmballageOpMenge(i) = 120: i = i + 1
LauerEmballagePzn(i) = 2182399: LauerEmballageOpName(i) = "VERSCHLUSS GL 18 KINDERGES": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2182407: LauerEmballageOpName(i) = "VERSCHLUSS GL 28 KINDERGES": LauerEmballageOpMenge(i) = 1: i = i + 1
LauerEmballagePzn(i) = 2599610: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 25: i = i + 1
LauerEmballagePzn(i) = 2599627: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 30: i = i + 1
LauerEmballagePzn(i) = 2599633: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 50: i = i + 1
LauerEmballagePzn(i) = 2599656: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 75: i = i + 1
LauerEmballagePzn(i) = 2599662: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 100: i = i + 1
LauerEmballagePzn(i) = 2599679: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 125: i = i + 1
LauerEmballagePzn(i) = 2599685: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 150: i = i + 1
LauerEmballagePzn(i) = 2599691: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 200: i = i + 1
LauerEmballagePzn(i) = 2599716: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 250: i = i + 1
LauerEmballagePzn(i) = 2599722: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 300: i = i + 1
LauerEmballagePzn(i) = 2599739: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL": LauerEmballageOpMenge(i) = 500: i = i + 1
LauerEmballagePzn(i) = 2599745: LauerEmballageOpName(i) = "WEITHALSGLAS BRAUN KOMPL.": LauerEmballageOpMenge(i) = 1000: i = i + 1

Dim OpEmballageName$()
Dim OpEmballageMenge#()
Dim OpEmballagePreis#()
iAnz = -1

Seek #ARBEMB%, (128&) + 1
Do
    Get #ARBEMB%, , ArbEmbRec
    If (EOF(ARBEMB%)) Then Exit Do
    If (ArbEmbRec.Typ = "E") Then
        h = Trim(iTrim(ArbEmbRec.kurz))
        Call OemToChar(h$, h$)
        
        a! = ArbEmbRec.menge
        Call DxToIEEEs(a!)
        
        kp = ArbEmbRec.kp
        Call DxToIEEEd(kp)
        kp = kp / 100#
        
        gef = -1
        For i = 0 To UBound(LauerEmballagePzn)
            If (h = LauerEmballageOpName(i)) And (a = LauerEmballageOpMenge(i)) Then
                gef = i
                Exit For
            End If
        Next i
        If (gef < 0) Then
            iAnz = iAnz + 1
            ReDim Preserve OpEmballageName(iAnz)
            ReDim Preserve OpEmballageMenge(iAnz)
            ReDim Preserve OpEmballagePreis(iAnz)
            OpEmballageName(iAnz) = h
            OpEmballageMenge(iAnz) = a
            OpEmballagePreis(iAnz) = kp
        End If

    End If
Loop
Close #ARBEMB%

Set MatchArt = New clsMatchArtikelADO
Call MatchArt.InitMatch

TaxmusterDBok = (Dir(TAXMUSTER_DB) <> "")
If (TaxmusterDBok) Then
    Set TaxmusterDB = OpenDatabase(TAXMUSTER_DB, False, False)

    SQLStr$ = "SELECT * FROM TaxmusterZeilen WHERE Flag=0"
    SQLStr$ = SQLStr$ + " ORDER BY Id"
    Set TaxmusterRec = TaxmusterDB.OpenRecordset(SQLStr$)
    Do
        If (TaxmusterRec.EOF) Then
            Exit Do
        End If
                
        h = Trim(CheckNullStr(TaxmusterRec!Name))
        a = CheckNullDouble(TaxmusterRec!ActMenge)
        
        gef = -1
        For i = 0 To UBound(LauerEmballagePzn)
            If (h = LauerEmballageOpName(i)) And (a = LauerEmballageOpMenge(i)) Then
                gef = i
                Exit For
            End If
        Next i
        If (gef >= 0) Then
            SQLStr = "UPDATE TaxMusterZeilen SET PZN =" + CStr(LauerEmballagePzn(gef))
            SQLStr = SQLStr + " WHERE Id=" + CStr(TaxmusterRec!Id)
            TaxmusterDB.Execute SQLStr
        Else
            h = Trim(CheckNullStr(TaxmusterRec!Name))
            a = CheckNullDouble(TaxmusterRec!ActMenge)
            kp = CheckNullDouble(TaxmusterRec!ActPreis)
            
            For i = 0 To UBound(OpEmballageName)
                If (h = OpEmballageName(i)) And (a = OpEmballageMenge(i)) Then
                    gef = i
                    Exit For
                End If
            Next i
            If (gef >= 0) Then
                kp = OpEmballagePreis(gef)
            End If
            
            SQLStr = "SELECT TOP 1 Id,PZN,Name,FloatMenge AS Menge,Einheit,Kp FROM Artikel AS A"
            SQLStr = SQLStr + " WHERE (LTRIM(RTRIM(Name))='" + h + "') AND (FloatMenge=" + uFormat(a, "0.00") + ")"
            SQLStr = SQLStr + " ORDER BY Menge"
            RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
            If (RezRbhLauerRec.EOF) Then
                sKurz$ = h + Format(a, "00000")
                Komprimiert = MatchArt.MakeMatchCode(sKurz$, 2)
                SortierName = MakeSortierName(Komprimiert)
                
                SQLStr = "INSERT INTO Artikel (Pzn,SortierName,Komprimiert)"
                SQLStr = SQLStr + " VALUES (" + CStr(0) + "," + "'" + SortierName + "'" + "," + "'" + Komprimiert + "'"
                SQLStr = SQLStr + ")"
                Call Hilfstaxe.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
                    
                Dim sRec As New ADODB.Recordset
                SQLStr = "SELECT @@Identity AS NewID"
                sRec.Open SQLStr, Hilfstaxe.ActiveConn
                If Not (sRec.EOF) Then
                    HSatz = sRec!NewId
                End If
                sRec.Close
                
                SQLStr = "UPDATE Artikel SET"
                SQLStr = SQLStr + " Name='" + h + "'"
                SQLStr = SQLStr + ", Menge='" + Left(Format(a, "0") + Space(5), 5) + "'"
                SQLStr = SQLStr + ", Einheit='" + Space(2) + "'"
                SQLStr = SQLStr + ", EK=" + uFormat(0, "0.00")
                SQLStr = SQLStr + ", VK=" + uFormat(0, "0.00")
                SQLStr = SQLStr + ", KP=" + uFormat(kp, "0.00")
                SQLStr = SQLStr + ", GSTUFE='" + Right$(Space$(7) + uFormat(a, "0.000"), 7) + "'"
                SQLStr = SQLStr + ", Emballage=" + CStr(1)
                SQLStr = SQLStr + ", FloatMenge=" + uFormat(a, "0.0000")
                SQLStr = SQLStr + " WHERE Id = " + CStr(HSatz)
                Call Hilfstaxe.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
            End If
            RezRbhLauerRec.Close
        End If

        TaxmusterRec.MoveNext
    Loop
    TaxmusterDB.Close
End If

On Error Resume Next
Kill "rezrbh.apo"
Name "rezrbh.dat" As "rezrbh.apo"
On Error GoTo DefErr

'APPLIKATOR UNGUATOR F. EMUL. 1
'BODENBEUTEL 250
'BODENBEUTEL 500
'BODENBEUTEL 1000
'BODENBEUTEL 2500
'BODENBEUTEL GEFUETTERT       375
'BODENBEUTEL GEFUETTERT       500
'BODENBEUTEL GEFUETTERT       1000
'DEOROLLER 50
'DEOROLLER 75
'DOSIERLOEFFEL 1, 7
'DOSIERLOEFFEL 5
'DOSIERSPRAYER 1
'DOSIERSPRÅHER 1
'FALTKARTON.SUPP.GIESSF.10    1
'FALTSCHACHTEL F. TUBEN GR. 9 1
'FALTSCHACHTEL F. TUBEN GR.7  1
'FALTSCHACHTEL TB. STAND. 120 1
'FALTSCHACHTEL TUBEN 100      1
'FALTSCHACHTEL TUBEN 20       1
'FALTSCHACHTEL TUBEN 30       1
'FALTSCHACHTEL TUBEN 50       1
'FLACHBEUTEL GR.  5           10
'FLACHBEUTEL GR.  7           15
'FLACHBEUTEL GR.  9           20
'FLACHBEUTEL GR. 11           50
'FLACHBEUTEL GR. 12           75
'FLACHBEUTEL GR. 13           100
'FLASCHE MIT SPATEAUFSATZ     10
'FLASCHE MIT SPATEAUFSATZ     20
'FLASCHE MIT SPATEAUFSATZ     30
'FLASCHE MIT SPATELAUFSATZ    50
'GELATINEWACHSKAPSEL 60
'GEWINDEFLASCHE GL 18         10
'GEWINDEFLASCHE GL 18         20
'GEWINDEFLASCHE GL 18         30
'GEWINDEFLASCHE GL 18         50
'GEWINDEFLASCHE GL 18         100
'GIESSFORM FUER 1 SUPP.       1
'GUMMIHOHLSTOPFEN 1
'HOHLSTOPFEN 21 MM            1
'INFUSIONSFLASCHE 125
'INFUSIONSFLASCHE 250
'INJEKTIONSFLASCHE 10
'INJEKTIONSSTOPFEN GR. 20     1
'KLYSMENFLASCHE 125
'KRUKE (KUNSTST.) DECKEL ROT  20
'KRUKE (KUNSTST.) DECKEL ROT  50
'KRUKE (KUNSTST.) DECKEL ROT  100
'KRUKE (KUNSTST.) DECKEL ROT  200
'KRUKE APOSAFE VERSCH. BODEN  20
'KRUKE APOSAFE VERSCH. BODEN  30
'KRUKE APOSAFE VERSCH. BODEN  50
'KRUKE APOSAFE VERSCH. BODEN  100
'KUNSTSTOFFFLASCHE 250
'KUNSTSTOFFFLASCHE 500
'KUNSTSTOFFFLASCHE 1000
'KUNSTSTOFFTUBE 50
'KUNSTSTOFFTUBE 100
'MEDIZINGLAS (BRAUN)          50
'MEDIZINGLAS (BRAUN)          100
'MEDIZINGLAS (BRAUN)          150
'MEDIZINGLAS (BRAUN)          200
'MEDIZINGLAS (BRAUN)          250
'MEDIZINGLAS (BRAUN)          300
'MEDIZINGLAS (BRAUN)          500
'MEDIZINGLAS (BRAUN)          1000
'METHADONFLASCHE 16
'METHADONFLASCHE 30
'MISCHSCHEIBE 1
'NASENSALBEN-APPLIKATOR       1
'OBLATENKAPSEL 0.5G GR. 0     1
'OBLATENKAPSEL 0.7G GR. 1     1
'OBLATENKAPSEL 0.8G GR. 2     1
'OBLATENKAPSEL KOMPL. 0.8 G   1
'PILLENGLAS 5
'PILLENGLAS 10
'PILLENGLAS 15
'PILLENGLAS 20
'PILLENGLAS 30
'PIPETTENGLAS ECKIG           10
'PIPETTENGLAS ECKIG           20
'PIPETTENGLAS ECKIG           30
'PIPETTENMONTUR GEPRUEFT      20
'PIPETTENMONTUR GEPRUEFT      50
'PULVERKAPSELGR. 3 POSTPAPIER 1
'PULVERKAPSELGR. 4 POSTPAPIER 1
'PULVERSCHACHTEL GR 10        400
'PULVERSCHACHTEL GR 2         20
'PULVERSCHACHTEL GR 3         30
'PULVERSCHACHTEL GR 4         50
'PULVERSCHACHTEL GR 6         100
'PULVERSCHACHTEL GR 7         175
'PULVERSCHACHTEL GR 8         250
'PULVERSCHACHTEL GR 9         300
'PULVERSCHACHTEL GR5          75
'QUALITAETSZUSCHLAG 10000
'RUNDFLASCHE HDPE             100
'RUNDFLASCHE HDPE             200
'RUNDFLASCHE LDPE             50
'SALBENDOSIERSPENDER 50
'SALBENDOSIERSPENDER 100
'SALBENDOSIERSPENDER 150
'SENKRECHTTROPFENMONTUR 1
'SICHERUNGSKAPPE 1
'SPATELFLASCHE OHNE UMKARTON  1
'SPATELMONTUR 20-50 ML        1
'SPRITZE 5
'SPRITZEINSATZ 100
'SPRITZEINSATZ 200
'STREUDOSE 100G               1
'STREUDOSE 50G                1
'SUPPOSITORIENKAESTCHEN 6
'SUPPOSITORIENKAESTCHEN 12 S  12
'TIEGEL 500                   25
'TOPITEC KRUKE                20
'TOPITEC KRUKE                30
'TOPITEC KRUKE                50
'TOPITEC KRUKE                100
'TOPITEC KRUKE                150
'TOPITEC KRUKE                200
'TROPFGLAS 75
'TROPFMONTUR 1
'TUBE (APONORM)               20
'TUBE (APONORM)               30
'TUBE (APONORM)               50
'TUBE (APONORM)               100
'TUBE APONORM                 25
'TUBE APONORM                 35
'TUBE MIT OLIVE GR 3          8
'TUBE MIT OLIVE GR 7          21
'WACHSKAPSEL GR.3             1
'WACHSKAPSEL GR.4             1
'WACHSKAPSEL GR.5             1
'WACHSKAPSEL GR.6             1
'ZERSTAEUBERPUMPE 100
'ZERSTAEUBERPUMPE DESINFECT.  1
'ZERSTAUBERPUMPE 100

Call DefErrPop
End Sub

Function MakeSortierName$(sKomprimiert$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeSortierName$")
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
Dim ret$
    
ret = sKomprimiert
    
SQLStr$ = "SELECT TOP 1 * FROM TAXE WHERE Komprimiert<'" + sKomprimiert + "'"
SQLStr = SQLStr + " ORDER BY Komprimiert DESC, SortierName DESC"
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If Not (TaxeRec.EOF) Then
    ret = TaxeRec!SortierName + "01"
End If
TaxeRec.Close

MakeSortierName = ret

Call DefErrPop
End Function

#If (WINREZX = 1) Then
#Else
Sub SubstitutionZuschlag()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SubstitutionZuschlag")
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
Dim i%, ActFlag%, AktMwSt%, FixAufschlagDa%, row%, aRow%, ArbeitDa%
Dim ActPreis#, ActMenge#, FaktorPreis#, SumGef‰ssPreis#, SumSpezPreis#
Dim h$, h2$
  
CannabisExtraktDichte = 0

SumPreis# = 0#
TeilPreis# = 0#
ProzPreis# = 0#
SumPreisZuz# = 0#
TeilPreisZuz# = 0#
ProzPreisZuz# = 0#
FaktorPreis# = 0#
TeilMenge = 0#
UnverarbeiteteAbgabe = 0

SumGef‰ssPreis# = 0#
SumSpezPreis# = 0#

AktMwSt% = para.Mwst(2)

frmSubstitutionsRezept.Show 1
If (SubstitutionsAbgaben(0) <= 0) Then
    Call DefErrPop: Exit Sub
End If

Dim SumSubstitutionsAbgaben As Integer
SumSubstitutionsAbgaben = 0
Dim SubstitutionsAbgabenStr As Integer
SubstitutionsAbgabenStr = 0
For i = 0 To 2
    SumSubstitutionsAbgaben = SumSubstitutionsAbgaben + SubstitutionsAbgaben(i)
    SubstitutionsAbgabenStr = (SubstitutionsAbgabenStr * 30) + SubstitutionsAbgaben(i)  'fr¸her * 10
Next i

Dim dSubstitutionsZuschlag As Double
Dim dEinzelZuschlag As Double
dSubstitutionsZuschlag = CalcSubstitutionZuschlag(SubstitutionsMg, SumSubstitutionsAbgaben, dEinzelZuschlag)
SubstitutionsEinzelpreis = dEinzelZuschlag

With frmTaxieren.flxTaxieren
'    For i% = 1 To (.Rows - 1)
'        If (InStr(UCase(.TextMatrix(i, 3)), "(ED:") > 0) Then
'            .RemoveItem (i)
'            Exit For
'        End If
'    Next
    For i% = 1 To (.Rows - 1)
        If (InStr(UCase(.TextMatrix(i, 3)), "E-DOSIS") > 0) Then
            .RemoveItem (i)
            Exit For
        End If
    Next
    For i% = 1 To (.Rows - 1)
        If (InStr(UCase(.TextMatrix(i, 3)), "FIX-AUFSCHLAG") > 0) Then
            .RemoveItem (i)
            Exit For
        End If
    Next
    For i% = 1 To (.Rows - 1)
        If (InStr(UCase(.TextMatrix(i, 3)), "BTM-GEB‹HR") > 0) Then
            .RemoveItem (i)
            Exit For
        End If
    Next
    For i% = 1 To (.Rows - 1)
        If (InStr(UCase(.TextMatrix(i, 3)), UCase("Honorierung des Sichtbezuges")) > 0) Then
            .RemoveItem (i)
            Exit For
        End If
    Next
    
    If (Trim(.TextMatrix(.Rows - 1, 0)) <> "") Then
        .AddItem " "
    End If
    
    .row = .Rows - 1
    row% = .row
    
    With TaxierRec
        .pzn = Space$(Len(.pzn))
        .kurz = Left$("Preis (E-Dosis:" + CStr(SubstitutionsMg) + ", Dosen:" + CStr(SumSubstitutionsAbgaben) + ")" + Space(Len(.kurz)), Len(.kurz))
        .menge = Space$(Len(.menge))
        .Meh = Space$(Len(.Meh))
        .kp = dEinzelZuschlag
        .GStufe = SubstitutionsMg   '0
        
        .ActMenge = SubstitutionsAbgabenStr '0#
        .ActPreis = dSubstitutionsZuschlag  ' 0
        
        .flag = MAG_PREISEINGABE
    End With
    Call frmTaxieren.ZeigeTaxierZeile(.row)

'        If (ParenteralRezept = 20) Or (ParenteralRezept = 22) Or (ParenteralRezept = 24) Then
'            .AddItem " "
'            .row = .Rows - 1
'            row% = .row
'
'            With TaxierRec
'                .pzn = Space$(Len(.pzn))
'                .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
'                .menge = Space$(Len(.menge))
'                .meh = Space$(Len(.meh))
'                .kp = 0
'                .Gstufe = 0
'
'                .ActMenge = 0#
'                .ActPreis = 8.35
'
'                .flag = MAG_PREISEINGABE
'            End With
'            Call frmTaxieren.ZeigeTaxierZeile(.row)
'        End If
        
    .AddItem " "
    .row = .Rows - 1
    row% = .row
    
    With TaxierRec
        .pzn = "02567001"   ' Space$(Len(.pzn))
        .kurz = Left$("BTM-Geb¸hr " + Space$(Len(.kurz)), Len(.kurz))
        .menge = Space$(Len(.menge))
        .Meh = Space$(Len(.Meh))
        .kp = 0
        .GStufe = 0
        
        .ActMenge = 1#
        .ActPreis = 3.58
        .kp = .ActPreis
        
        .flag = MAG_PREISEINGABE
    End With
    Call frmTaxieren.ZeigeTaxierZeile(.row)
    
    Dim SumSubstitutionsAbgabenSichtbezug As Integer
    h = MyInputBox("Honorierung des Sichtbezuges (Anzahl): ", "Sichtbezug der Opioidsubstitution", "")
    SumSubstitutionsAbgabenSichtbezug = Val(h)

    If (SumSubstitutionsAbgabenSichtbezug > 0) Then
        .AddItem " "
        .row = .Rows - 1
        row% = .row

        With TaxierRec
            .pzn = "18774506"   ' Space$(Len(.pzn))
            .kurz = Left$("Honorierung des Sichtbezuges der Opioidsubstitution " + Space$(100), 100) '+ Space$(Len(.kurz)), Len(.kurz))

            .menge = Space$(Len(.menge))
            .Meh = Space$(Len(.Meh))
            .kp = 0
            .GStufe = 0
            
            .ActMenge = SumSubstitutionsAbgabenSichtbezug
            .ActPreis = 5.7 * SumSubstitutionsAbgabenSichtbezug
            .kp = .ActPreis

            .flag = MAG_PREISEINGABE
        End With
        Call frmTaxieren.ZeigeTaxierZeile(.row)
    End If

End With
    
Call DefErrPop
End Sub
#End If

Function CalcSubstitutionZuschlag#(dEinzeldosisMg#, iSumEinzeldosis%, dEinzelZuschlag#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcSubstitutionZuschlag#")
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
Dim i%, ind%
Dim dret#
Dim sPreise$

dret = 0
sPreise = ""
dEinzelZuschlag = 0

If (ParenteralRezept = 16) Or (ParenteralRezept = 17) Then
    If (dEinzeldosisMg <= 10) Then
        dret = 2.08
    ElseIf (dEinzeldosisMg <= 20) Then
        dret = 2.11
    ElseIf (dEinzeldosisMg <= 30) Then
        dret = 2.14
    ElseIf (dEinzeldosisMg <= 40) Then
        dret = 2.17
    ElseIf (dEinzeldosisMg <= 50) Then
        dret = 2.19
    ElseIf (dEinzeldosisMg <= 60) Then
        dret = 2.22
    ElseIf (dEinzeldosisMg <= 70) Then
        dret = 2.25
    ElseIf (dEinzeldosisMg <= 80) Then
        dret = 2.27
    ElseIf (dEinzeldosisMg <= 90) Then
        dret = 2.3
    ElseIf (dEinzeldosisMg <= 100) Then
        dret = 2.34
    ElseIf (dEinzeldosisMg <= 110) Then
        dret = 2.37
    ElseIf (dEinzeldosisMg <= 120) Then
        dret = 2.39
    ElseIf (dEinzeldosisMg <= 130) Then
        dret = 2.42
    ElseIf (dEinzeldosisMg <= 140) Then
        dret = 2.45
    ElseIf (dEinzeldosisMg <= 150) Then
        dret = 2.48
    ElseIf (dEinzeldosisMg <= 160) Then
        dret = 2.85
    ElseIf (dEinzeldosisMg <= 170) Then
        dret = 2.88
    ElseIf (dEinzeldosisMg <= 180) Then
        dret = 2.91
    ElseIf (dEinzeldosisMg <= 190) Then
        dret = 2.93
    ElseIf (dEinzeldosisMg <= 200) Then
        dret = 2.95
    Else
        dret = 2.95 + ((dEinzeldosisMg - 200) \ 10 + 1) * 0.03
    End If
ElseIf (ParenteralRezept = 18) Or (ParenteralRezept = 19) Then
    If (dEinzeldosisMg <= 5) Then
        dret = 2.54
    ElseIf (dEinzeldosisMg <= 7.5) Then
        dret = 2.63
    ElseIf (dEinzeldosisMg <= 10) Then
        dret = 2.72
    ElseIf (dEinzeldosisMg <= 12.5) Then
        dret = 2.81
    ElseIf (dEinzeldosisMg <= 15) Then
        dret = 2.9
    ElseIf (dEinzeldosisMg <= 17.5) Then
        dret = 2.98
    ElseIf (dEinzeldosisMg <= 20) Then
        dret = 3.07
    ElseIf (dEinzeldosisMg <= 22.5) Then
        dret = 3.16
    ElseIf (dEinzeldosisMg <= 25) Then
        dret = 3.25
    ElseIf (dEinzeldosisMg <= 27.5) Then
        dret = 3.33
    ElseIf (dEinzeldosisMg <= 30) Then
        dret = 3.43
    ElseIf (dEinzeldosisMg <= 32.5) Then
        dret = 3.51
    ElseIf (dEinzeldosisMg <= 35) Then
        dret = 3.6
    ElseIf (dEinzeldosisMg <= 37.5) Then
        dret = 3.69
    ElseIf (dEinzeldosisMg <= 40) Then
        dret = 3.77
    ElseIf (dEinzeldosisMg <= 42.5) Then
        dret = 3.86
    ElseIf (dEinzeldosisMg <= 45) Then
        dret = 3.95
    ElseIf (dEinzeldosisMg <= 47.5) Then
        dret = 4.04
    ElseIf (dEinzeldosisMg <= 50) Then
        dret = 4.12
    ElseIf (dEinzeldosisMg <= 52.5) Then
        dret = 4.21
    ElseIf (dEinzeldosisMg <= 55) Then
        dret = 4.3
    ElseIf (dEinzeldosisMg <= 57.5) Then
        dret = 4.39
    ElseIf (dEinzeldosisMg <= 60) Then
        dret = 4.48
    ElseIf (dEinzeldosisMg <= 62.5) Then
        dret = 4.56
    ElseIf (dEinzeldosisMg <= 65) Then
        dret = 4.65
    ElseIf (dEinzeldosisMg <= 67.5) Then
        dret = 4.74
    ElseIf (dEinzeldosisMg <= 70) Then
        dret = 4.83
    ElseIf (dEinzeldosisMg <= 72.5) Then
        dret = 4.91
    ElseIf (dEinzeldosisMg <= 75) Then
        dret = 5
    Else
        dret = 5 + ((dEinzeldosisMg - 75) * 10 \ 25 + 1) * 0.09
    End If
ElseIf (ParenteralRezept = 20) Then
    If (dEinzeldosisMg <= 2) Then
        sPreise = "1.10;2.21;3.31;4.42;5.52;6.62;7.73;"
    ElseIf (dEinzeldosisMg <= 2.4) Then
        sPreise = "1.59;3.19;4.78;6.37;7.97;9.56;11,16;"
    ElseIf (dEinzeldosisMg <= 2.8) Then
        sPreise = "2.09;4.18;6.27;8.37;10.46;12.55;14.64;"
    ElseIf (dEinzeldosisMg <= 3.2) Then
        sPreise = "2.58;5.16;7.74;10.33;12.91;15.49;18.07;"
    ElseIf (dEinzeldosisMg <= 3.6) Then
        sPreise = "3.07;6.14;9.21;12.28;15.36;18.43;21.50;"
    ElseIf (dEinzeldosisMg <= 4) Then
        sPreise = "2.08;4.17;6.25;8.33;10.42;12.50;14.58;"
    ElseIf (dEinzeldosisMg <= 4.4) Then
        sPreise = "2.57;5.15;7.72;10.29;12.87;15.44;18.01;"
    ElseIf (dEinzeldosisMg <= 4.8) Then
        sPreise = "3.07;6.14;9.21;12.28;15.36;18.43;21.50;"
    ElseIf (dEinzeldosisMg <= 5.2) Then
        sPreise = "3.56;7.12;10.68;14.24;17.80;21.36;24.92;"
    ElseIf (dEinzeldosisMg <= 5.6) Then
        sPreise = "4.05;8.10;12.15;16.20;20.25;24.30;28.35;"
    ElseIf (dEinzeldosisMg <= 6) Then
        sPreise = "3.06;6.13;9.19;12.25;15.31;18.38;21.44;"
    ElseIf (dEinzeldosisMg <= 6.4) Then
        sPreise = "3.55;7.10;10.66;14.21;17.76;21.31;24.87;"
    ElseIf (dEinzeldosisMg <= 6.8) Then
        sPreise = "4.05;8.10;12.15;16.20;20.25;24.30;28.35;"
    ElseIf (dEinzeldosisMg <= 7.2) Then
        sPreise = "4.54;9.08;13.62;18.16;22.70;27.24;31.78;"
    ElseIf (dEinzeldosisMg <= 7.6) Then
        sPreise = "5.03;10.06;15.09;20.12;25.12;30.18;35.21;"
    ElseIf (dEinzeldosisMg <= 8) Then
        sPreise = "3.99;7.98;11.98;15.97;19.96;23.95;27.95;"
    ElseIf (dEinzeldosisMg <= 8.4) Then
        sPreise = "4.48;8.96;13.45;17.93;22.41;26.89;31.37;"
    ElseIf (dEinzeldosisMg <= 8.8) Then
        sPreise = "4.98;9.96;14.94;19.92;24.90;29.88;34.86;"
    ElseIf (dEinzeldosisMg <= 9.2) Then
        sPreise = "5.47;10.94;16.41;21.88;27.35;32.82;38.29;"
    ElseIf (dEinzeldosisMg <= 9.6) Then
        sPreise = "5.96;11.92;17.88;23.84;29.80;35.76;41.72;"
    ElseIf (dEinzeldosisMg <= 10) Then
        sPreise = "4.97;9.94;14.92;19.89;24.86;29.83;34.80;"
    ElseIf (dEinzeldosisMg <= 10.4) Then
        sPreise = "5.46;10.92;16.38;21.85;27.31;32.77;38.23;"
    ElseIf (dEinzeldosisMg <= 10.8) Then
        sPreise = "5.96;11.92;17.88;23.84;29.80;35.76;41.72;"
    ElseIf (dEinzeldosisMg <= 11.2) Then
        sPreise = "6.45;12.90;19.35;25.80;32.25;38.69;45.14;"
    ElseIf (dEinzeldosisMg <= 11.6) Then
        sPreise = "6.94;13.88;20.82;27.76;34.69;41.63;48.57;"
    ElseIf (dEinzeldosisMg <= 12) Then
        sPreise = "5.95;11.90;17.85;23.80;29.76;35.71;41.66;"
    ElseIf (dEinzeldosisMg <= 12.4) Then
        sPreise = "6.44;12.88;19.32;25.76;32.20;38.64;45.09;"
    ElseIf (dEinzeldosisMg <= 12.8) Then
        sPreise = "6.94;13.88;20.82;27.76;34.69;41.63;48.57;"
    ElseIf (dEinzeldosisMg <= 13.2) Then
        sPreise = "7.43;14.86;22.29;29.71;37.14;44.57;52.00;"
    ElseIf (dEinzeldosisMg <= 13.6) Then
        sPreise = "7.92;15.84;23.75;31.67;39.59;47.51;55.43;"
    ElseIf (dEinzeldosisMg <= 14) Then
        sPreise = "6.93;13.86;20.79;27.72;34.65;41.58;48.51;"
    ElseIf (dEinzeldosisMg <= 14.4) Then
        sPreise = "7.42;14.84;22.26;29.68;37.10;44.52;51.94;"
    ElseIf (dEinzeldosisMg <= 14.8) Then
        sPreise = "7.92;15.84;23.75;31.67;39.59;47.51;55.43;"
    ElseIf (dEinzeldosisMg <= 15.2) Then
        sPreise = "8.41;16.82;25.22;33.63;42.04;50.45;58.86;"
    ElseIf (dEinzeldosisMg <= 15.6) Then
        sPreise = "8.90;17.80;26.69;35.59;44.49;53.39;62.28;"
    ElseIf (dEinzeldosisMg <= 16) Then
        sPreise = "7.86;15.72;23.58;31.44;39.30;47.16;55.02;"
    ElseIf (dEinzeldosisMg <= 16.4) Then
        sPreise = "8.35;16.70;25.05;33.40;41.75;50.10;58.45;"
    ElseIf (dEinzeldosisMg <= 16.8) Then
        sPreise = "8.85;17.70:26.54;35.39;44.24;53.09;61.93;"
    ElseIf (dEinzeldosisMg <= 17.2) Then
        sPreise = "9.34;18.68;28.01;37.35;46.69;56.03;65.36;"
    ElseIf (dEinzeldosisMg <= 17.6) Then
        sPreise = "9.83;19.65;29.48;39.31;49.14;58.96;68.79;"
    ElseIf (dEinzeldosisMg <= 18) Then
        sPreise = "8.84;17.68;26.52;35.36;44.20;53.04;61.88;"
    ElseIf (dEinzeldosisMg <= 18.4) Then
        sPreise = "9.33;18.66;27.99;37.32;46.65;55.98;65.30;"
    ElseIf (dEinzeldosisMg <= 18.8) Then
        sPreise = "9.83;19.65;29.48;39.31;49.14;58.96;68.79;"
    ElseIf (dEinzeldosisMg <= 19.2) Then
        sPreise = "10.32;20.63;30.95;41.27;51.58;61.90;72.22;"
    ElseIf (dEinzeldosisMg <= 19.6) Then
        sPreise = "10.81;21.61;32.42;43.23;54.03;64.84;75.65;"
    ElseIf (dEinzeldosisMg <= 20) Then
        sPreise = "9.82;19.64;29.46;39.28;49.09;58.91;68.73;"
    ElseIf (dEinzeldosisMg <= 20.4) Then
        sPreise = "10.31;20.62;30.93;41.23;51.54;61.85;72.16;"
    ElseIf (dEinzeldosisMg <= 20.8) Then
        sPreise = "10.81;21.61;32.42;43.23;54.03;64.84;75.65;"
    ElseIf (dEinzeldosisMg <= 21.2) Then
        sPreise = "11.30;22.59;33.89;45.19;56.48;67.78;79.07;"
    ElseIf (dEinzeldosisMg <= 21.6) Then
        sPreise = "11.79;23.57;35.36;47.14;58.93;70.72;82.50;"
    ElseIf (dEinzeldosisMg <= 22) Then
        sPreise = "10.80;21.60;32.39;43.19;53.99;64.79;75.59;"
    ElseIf (dEinzeldosisMg <= 22.4) Then
        sPreise = "11.29;22.58;33.86;45.15;56.44;67.73;79.02;"
    ElseIf (dEinzeldosisMg <= 22.8) Then
        sPreise = "11.79;23.57;35.36;47.14;58.93;70.72;82.50;"
    ElseIf (dEinzeldosisMg <= 23.2) Then
        sPreise = "12.28;24.55;36.83;49.10;61.38;73.65;85.93;"
    ElseIf (dEinzeldosisMg <= 23.6) Then
        sPreise = "12.77;25.53;38.30;51.06;63.83;76.59;89.36;"
    ElseIf (dEinzeldosisMg <= 24) Then
        sPreise = "11.73;23.46;35.18;46.91;58.64;70.37;82.10;"
    Else
    End If
ElseIf (ParenteralRezept = 21) Then
    If (dEinzeldosisMg <= 2) Then
        sPreise = "1.33;2.66;3.99;5.32;6.65;7.98;9.31;"
    ElseIf (dEinzeldosisMg <= 4) Then
        sPreise = "2.51;5.02;7.53;10.04;12.55;15.06;17.57;"
    ElseIf (dEinzeldosisMg <= 6) Then
        sPreise = "3.69;7.38;11.07;14.76;18.45;22.14;25.83;"
    ElseIf (dEinzeldosisMg <= 8) Then
        sPreise = "4.81;9.62;14.43;19.24;24.05;28.86;33.67;"
    ElseIf (dEinzeldosisMg <= 10) Then
        sPreise = "5.99;11.98;17.97;23.96;29.95;35.94;41.93;"
    ElseIf (dEinzeldosisMg <= 12) Then
        sPreise = "7.17;14.34;21.51;28.68;35.85;43.02;50.19;"
    ElseIf (dEinzeldosisMg <= 14) Then
        sPreise = "8.35;16.70;25.05;33.40;41.75;50.10;58.45;"
    ElseIf (dEinzeldosisMg <= 16) Then
        sPreise = "9.47;18.94;28.41;37.88;47.35;56.82;66.29;"
    ElseIf (dEinzeldosisMg <= 18) Then
        sPreise = "10.65;21.30;31.95;42.60;53.25;63.90;74.55;"
    ElseIf (dEinzeldosisMg <= 20) Then
        sPreise = "11.83;23.66;35.49;47.32;59.15;70.98;82.81;"
    ElseIf (dEinzeldosisMg <= 22) Then
        sPreise = "13.01;26.02;39.03;52.04;65.05;78.06;91.07;"
    ElseIf (dEinzeldosisMg <= 24) Then
        sPreise = "14.13;28.26;42.39;56.52;70.65;84,78;98.91;"
    Else
    End If
ElseIf (ParenteralRezept = 22) Then
    If (dEinzeldosisMg <= 2) Then
        sPreise = "1.58;3.17;4.75;6.33;7.91;9.50;11.08;"
    ElseIf (dEinzeldosisMg <= 2.4) Then
        sPreise = "2.28;4.57;6.85;9.14;11.42;13.71;15.99;"
    ElseIf (dEinzeldosisMg <= 2.8) Then
        sPreise = "3;6;9;12;14.99;17.99;20.99;"
    ElseIf (dEinzeldosisMg <= 3.2) Then
        sPreise = "3.70;7.40;11.10;14.80;18.50;22.21;25.91;"
    ElseIf (dEinzeldosisMg <= 3.6) Then
        sPreise = "4.40;8.81;13.21;17.61;22.02;26.42;30.82;"
    ElseIf (dEinzeldosisMg <= 4) Then
        sPreise = "2.99;5.97;8.96;11.95;14.93;17.92;20.91;"
    ElseIf (dEinzeldosisMg <= 4.4) Then
        sPreise = "3.69;7.38;11.07;14.76;18.45;22.13;25.82;"
    ElseIf (dEinzeldosisMg <= 4.8) Then
        sPreise = "4.40;8.81;13.21;17.61;22.02;26.42;30.82;"
    ElseIf (dEinzeldosisMg <= 5.2) Then
        sPreise = "5.11;10.21;15.32;20.42;25.53;30.63;35.74;"
    ElseIf (dEinzeldosisMg <= 5.6) Then
        sPreise = "5.81;11.61;17.42;23.23;29.04;34.84;40.65;"
    ElseIf (dEinzeldosisMg <= 6) Then
        sPreise = "4.39;8.78;13.17;17.56;21.96;26.35;30.74;"
    ElseIf (dEinzeldosisMg <= 6.4) Then
        sPreise = "5.09;10.19;15.28;20.37;25.47;30.56;35.65;"
    ElseIf (dEinzeldosisMg <= 6.8) Then
        sPreise = "5.81;11.61;17.42;23.23;29.04;34.84;40.65;"
    ElseIf (dEinzeldosisMg <= 7.2) Then
        sPreise = "6.51;13.02;19.53;26.04;32.55;39.06;45.57;"
    ElseIf (dEinzeldosisMg <= 7.6) Then
        sPreise = "7.21;14.42;21.63;28.85;36.06;43.27;50.48;"
    ElseIf (dEinzeldosisMg <= 8) Then
        sPreise = "5.72;11.45;17.17;22.90;28.62;34.34;40.07;"
    ElseIf (dEinzeldosisMg <= 8.4) Then
        sPreise = "6.43;12.85;19.28;25.70;32.13;38.56;44.98;"
    ElseIf (dEinzeldosisMg <= 8.8) Then
        sPreise = "7.14;14.28;21.42;28.56;35.70;42.84;49.98;"
    ElseIf (dEinzeldosisMg <= 9.2) Then
        sPreise = "7.84;15.68;23.53;31.37;39.21;47.05;54.89;"
    ElseIf (dEinzeldosisMg <= 9.6) Then
        sPreise = "8.54;17.09;25.63;34.18;42.72;51.27;59.81;"
    ElseIf (dEinzeldosisMg <= 10) Then
        sPreise = "7.13;14.26;21.38;28.51;35.64;42.77;49.90;"
    ElseIf (dEinzeldosisMg <= 10.4) Then
        sPreise = "7.83;15.66;23.49;31.32;39.15;46.98;54.81;"
    ElseIf (dEinzeldosisMg <= 10.8) Then
        sPreise = "8.54;17.09;25.63;34.18;42.72;51.27;59.81;"
    ElseIf (dEinzeldosisMg <= 11.2) Then
        sPreise = "9.25;18,49;27.74;36.99;46.23;55.48;64.72;"
    ElseIf (dEinzeldosisMg <= 11.6) Then
        sPreise = "9.95;19.90;29.85;39.79;49.74;59.69;69.64;"
    ElseIf (dEinzeldosisMg <= 12) Then
        sPreise = "8.53;17.06;25.60;34.13;42.66;51.19;59.73;"
    ElseIf (dEinzeldosisMg <= 12.4) Then
        sPreise = "9.23;18.47;27.70;36.94;46.17;55.41;64.64;"
    ElseIf (dEinzeldosisMg <= 12.8) Then
        sPreise = "9.95;19.90;29.85;39.79;49.74;59.69;69.64;"
    ElseIf (dEinzeldosisMg <= 13.2) Then
        sPreise = "10.65;21.30;31.95;42.60;53.25;63.90;74.55;"
    ElseIf (dEinzeldosisMg <= 13.6) Then
        sPreise = "11.35;22.71;34.06;45.41;56.76;68.12;79.47;"
    ElseIf (dEinzeldosisMg <= 14) Then
        sPreise = "9.94;19.87;29.81;39.75;49.68;59.62;69.56;"
    ElseIf (dEinzeldosisMg <= 14.4) Then
        sPreise = "10.64;21.28;31.92;42.55;53.19;63.83;74.47;"
    ElseIf (dEinzeldosisMg <= 14.8) Then
        sPreise = "11.35;22.71;34.06;45.41;56.76;68.12;79.47;"
    ElseIf (dEinzeldosisMg <= 15.2) Then
        sPreise = "12.05;24.11;36.16;48.22;60.27;72.33;84.38;"
    ElseIf (dEinzeldosisMg <= 15.6) Then
        sPreise = "12.76;25.51;38.27;51.03;63.78;76.54;89.30;"
    ElseIf (dEinzeldosisMg <= 16) Then
        sPreise = "11.27;22.54;33.81;45.08;56.35;67.62;78.89;"
    ElseIf (dEinzeldosisMg <= 16.4) Then
        sPreise = "11.97;23.94;35.91;47.89;59.86;71.83;83.80;"
    ElseIf (dEinzeldosisMg <= 16.8) Then
        sPreise = "12.69;25.37;38.06;50.74;63.43;76.11;88.80;"
    ElseIf (dEinzeldosisMg <= 17.2) Then
        sPreise = "13.39;26.78;40.16;53.55;66.94;80.33;93.71;"
    ElseIf (dEinzeldosisMg <= 17.6) Then
        sPreise = "14.09;28.18;42.27;56.36;70.45;84.54;98.63;"
    ElseIf (dEinzeldosisMg <= 18) Then
        sPreise = "12.67;25.35;38.02;50.69;63.37;76.04;88.71;"
    ElseIf (dEinzeldosisMg <= 18.4) Then
        sPreise = "13.38;26.75;40.13;53.50;66.88;80.25;93.63;"
    ElseIf (dEinzeldosisMg <= 18.8) Then
        sPreise = "14.09;28.18;42.27;56.36;70.45;84.54;98.63;"
    ElseIf (dEinzeldosisMg <= 19.2) Then
        sPreise = "14.79;29.58;44.38;59.17;73.96;88.75;103.54;"
    ElseIf (dEinzeldosisMg <= 19.6) Then
        sPreise = "15.49;30.99;46.48;61.98;77.47;92.96;108.46;"
    ElseIf (dEinzeldosisMg <= 20) Then
        sPreise = "14.08;28.16;42.23;56.31;70.39;84.47;98.54;"
    ElseIf (dEinzeldosisMg <= 20.4) Then
        sPreise = "14.78;29.56;44.34;59.12;73.90;88.68;103.46;"
    ElseIf (dEinzeldosisMg <= 20.8) Then
        sPreise = "15.49;30.99;46.48;61.98;77.47;92.96;108.46;"
    ElseIf (dEinzeldosisMg <= 21.2) Then
        sPreise = "16.20;32.39;48.59;64.78;80.98;97.18;113.37;"
    ElseIf (dEinzeldosisMg <= 21.6) Then
        sPreise = "16.90;33.80;50.69;67.59;84.49;101.39;118.29;"
    ElseIf (dEinzeldosisMg <= 22) Then
        sPreise = "15.48;30.96;46.45;61.93;77.41;92.89;108.37;"
    ElseIf (dEinzeldosisMg <= 22.4) Then
        sPreise = "16.18;32.37;48.55;64.74;80.92;97.10;113.29;"
    ElseIf (dEinzeldosisMg <= 22.8) Then
        sPreise = "16.90;33.80;50.69;67.59;84.49;101.39;118.29;"
    ElseIf (dEinzeldosisMg <= 23.2) Then
        sPreise = "17.60;35.20;52.80;70.40;88.00;105.60;123.20;"
    ElseIf (dEinzeldosisMg <= 23.6) Then
        sPreise = "18.30;36.60;54.91;73.21;91.51;109.81;128.12;"
    ElseIf (dEinzeldosisMg <= 24) Then
        sPreise = "16.81;33.63;50.44;67.26;84.07;100.89;117.70;"
    Else
    End If
ElseIf (ParenteralRezept = 23) Then
    Dim dPreis#
    If (dEinzeldosisMg <= 2) Then
        dPreis = 1.33
    ElseIf (dEinzeldosisMg <= 2.4) Then
        dPreis = 1.92
    ElseIf (dEinzeldosisMg <= 2.8) Then
        dPreis = 2.52
    ElseIf (dEinzeldosisMg <= 3.2) Then
        dPreis = 3.11
    ElseIf (dEinzeldosisMg <= 3.6) Then
        dPreis = 3.7
    ElseIf (dEinzeldosisMg <= 4) Then
        dPreis = 2.51
    ElseIf (dEinzeldosisMg <= 4.4) Then
        dPreis = 3.1
    ElseIf (dEinzeldosisMg <= 4.8) Then
        dPreis = 3.7
    ElseIf (dEinzeldosisMg <= 5.2) Then
        dPreis = 4.29
    ElseIf (dEinzeldosisMg <= 5.6) Then
        dPreis = 4.88
    ElseIf (dEinzeldosisMg <= 6) Then
        dPreis = 3.69
    ElseIf (dEinzeldosisMg <= 6.4) Then
        dPreis = 4.28
    ElseIf (dEinzeldosisMg <= 6.8) Then
        dPreis = 4.88
    ElseIf (dEinzeldosisMg <= 7.2) Then
        dPreis = 5.47
    ElseIf (dEinzeldosisMg <= 7.6) Then
        dPreis = 6.06
    ElseIf (dEinzeldosisMg <= 8) Then
        dPreis = 4.81
    ElseIf (dEinzeldosisMg <= 8.4) Then
        dPreis = 5.4
    ElseIf (dEinzeldosisMg <= 8.8) Then
        dPreis = 6
    ElseIf (dEinzeldosisMg <= 9.2) Then
        dPreis = 6.59
    ElseIf (dEinzeldosisMg <= 9.6) Then
        dPreis = 7.18
    ElseIf (dEinzeldosisMg <= 10) Then
        dPreis = 5.99
    ElseIf (dEinzeldosisMg <= 10.4) Then
        dPreis = 6.58
    ElseIf (dEinzeldosisMg <= 10.8) Then
        dPreis = 7.18
    ElseIf (dEinzeldosisMg <= 11.2) Then
        dPreis = 7.77
    ElseIf (dEinzeldosisMg <= 11.6) Then
        dPreis = 8.36
    ElseIf (dEinzeldosisMg <= 12) Then
        dPreis = 7.17
    ElseIf (dEinzeldosisMg <= 12.4) Then
        dPreis = 7.76
    ElseIf (dEinzeldosisMg <= 12.8) Then
        dPreis = 8.36
    ElseIf (dEinzeldosisMg <= 13.2) Then
        dPreis = 8.95
    ElseIf (dEinzeldosisMg <= 13.6) Then
        dPreis = 9.54
    ElseIf (dEinzeldosisMg <= 14) Then
        dPreis = 8.35
    ElseIf (dEinzeldosisMg <= 14.4) Then
        dPreis = 8.94
    ElseIf (dEinzeldosisMg <= 14.8) Then
        dPreis = 9.54
    ElseIf (dEinzeldosisMg <= 15.2) Then
        dPreis = 10.13
    ElseIf (dEinzeldosisMg <= 15.6) Then
        dPreis = 10.72
    ElseIf (dEinzeldosisMg <= 16) Then
        dPreis = 9.47
    ElseIf (dEinzeldosisMg <= 16.4) Then
        dPreis = 10.06
    ElseIf (dEinzeldosisMg <= 16.8) Then
        dPreis = 10.66
    ElseIf (dEinzeldosisMg <= 17.2) Then
        dPreis = 11.25
    ElseIf (dEinzeldosisMg <= 17.6) Then
        dPreis = 11.84
    ElseIf (dEinzeldosisMg <= 18) Then
        dPreis = 10.65
    ElseIf (dEinzeldosisMg <= 18.4) Then
        dPreis = 11.24
    ElseIf (dEinzeldosisMg <= 18.8) Then
        dPreis = 11.84
    ElseIf (dEinzeldosisMg <= 19.2) Then
        dPreis = 12.43
    ElseIf (dEinzeldosisMg <= 19.6) Then
        dPreis = 13.02
    ElseIf (dEinzeldosisMg <= 20) Then
        dPreis = 11.83
    ElseIf (dEinzeldosisMg <= 20.4) Then
        dPreis = 12.42
    ElseIf (dEinzeldosisMg <= 20.8) Then
        dPreis = 13.02
    ElseIf (dEinzeldosisMg <= 21.2) Then
        dPreis = 13.61
    ElseIf (dEinzeldosisMg <= 21.6) Then
        dPreis = 14.2
    ElseIf (dEinzeldosisMg <= 22) Then
        dPreis = 13.01
    ElseIf (dEinzeldosisMg <= 22.4) Then
        dPreis = 13.6
    ElseIf (dEinzeldosisMg <= 22.8) Then
        dPreis = 14.2
    ElseIf (dEinzeldosisMg <= 23.2) Then
        dPreis = 14.79
    ElseIf (dEinzeldosisMg <= 23.6) Then
        dPreis = 15.38
    ElseIf (dEinzeldosisMg <= 24) Then
        dPreis = 14.13
    Else
    End If
    sPreise = ""
    For i = 1 To 7
        sPreise = sPreise + Format(dPreis * i, "0.00") + ";"
    Next i
End If


If (ParenteralRezept <= 19) Then
    dEinzelZuschlag = dret
    dret = dret * iSumEinzeldosis
ElseIf (sPreise <> "") Then
    ind = InStr(sPreise, ";")
    If (ind > 0) Then
        dEinzelZuschlag = xVal(Left(sPreise, ind - 1))
    End If
    
    If (iSumEinzeldosis > 7) Then
        dret = dEinzelZuschlag * iSumEinzeldosis
    Else
        For i = 1 To (iSumEinzeldosis - 1)
            ind = InStr(sPreise, ";")
            If (ind > 0) Then
                sPreise = Mid(sPreise, ind + 1)
            Else
                Exit For
            End If
        Next i
        ind = InStr(sPreise, ";")
        If (ind > 0) Then
            dret = xVal(Left(sPreise, ind - 1))
        End If
    End If
    If (ParenteralRezept = 22) Then
        dret = dret / 1.19
    End If
End If

CalcSubstitutionZuschlag = dret
    
Call DefErrPop
End Function


Sub CannabisZuschlag()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CannabisZuschlag")
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
Dim i%, ActFlag%, AktMwSt%, FixAufschlagDa%, row%, aRow%, ArbeitDa%
Dim ActPreis#, ActMenge#, FaktorPreis#, SumGef‰ssPreis#, SumSpezPreis#
Dim h$, h2$
  
CannabisExtraktDichte = 0

SumPreis# = 0#
TeilPreis# = 0#
ProzPreis# = 0#
SumPreisZuz# = 0#
TeilPreisZuz# = 0#
ProzPreisZuz# = 0#
FaktorPreis# = 0#
TeilMenge = 0#
UnverarbeiteteAbgabe = 0

SumGef‰ssPreis# = 0#
SumSpezPreis# = 0#

AktMwSt% = para.Mwst(2)

With frmTaxieren.flxTaxieren
    If (ParenteralRezept > 15) Then
        TeilMenge = 0
        For i% = 1 To (.Rows - 1)
            ActMenge# = iCDbl(.TextMatrix(i%, 1))
            ActFlag% = Val(.TextMatrix(i%, 6))
            
            If (ActFlag% < MAG_NN) Then
                If (ActFlag% <> MAG_GEFAESS) And (ActFlag% <> MAG_ARBEIT) And (ActFlag% <> MAG_SONSTIGES) And (ActFlag% <> MAG_PREISEINGABE) Then
                    If ((ActFlag% <> MAG_SPEZIALITAET) And (ActFlag% <> MAG_ANTEILIG)) Or (.TextMatrix(i%, 2) <> "ST") Then
                        TeilMenge# = TeilMenge# + ActMenge#
                    End If
                End If
            End If
            
            'ab 5.1.55 210926
            If (ParenteralRezept = 24) Or (ParenteralRezept = 25) Or (ParenteralRezept = 26) Or (ParenteralRezept = 27) Then
                If (ActFlag% = MAG_HILFSTAXE) Or (ActFlag% = MAG_SPEZIALITAET) Then
                    .TextMatrix(i%, 0) = Format(0, "0.00")
                    .TextMatrix(i%, 13) = Format(0, "0.00")
                End If
            End If
        Next
        If (TeilMenge > 0) Then
            For i% = 1 To (.Rows - 1)
                If (InStr(UCase(.TextMatrix(i, 3)), "FIXZUSCHLAG") > 0) Then
                    .RemoveItem (i)
                    Exit For
                End If
            Next
            For i% = 1 To (.Rows - 1)
                If (InStr(UCase(.TextMatrix(i, 3)), "FIX-AUFSCHLAG") > 0) Then
                    .RemoveItem (i)
                    Exit For
                End If
            Next
            For i% = 1 To (.Rows - 1)
                If (InStr(UCase(.TextMatrix(i, 3)), "BTM-GEB‹HR") > 0) Then
                    .RemoveItem (i)
                    Exit For
                End If
            Next
            
            If (ParenteralRezept = 31) Then
            Else
                If (Trim(.TextMatrix(.Rows - 1, 0)) <> "") Then
                    .AddItem " "
                End If
                
                .row = .Rows - 1
                row% = .row
                
                With TaxierRec
                    .pzn = Space$(Len(.pzn))
                    .kurz = Left$("Fixzuschlag" + Space$(Len(.kurz)), Len(.kurz))
                    .menge = Space$(Len(.menge))
                    .Meh = Space$(Len(.Meh))
                    .kp = 0
                    .GStufe = 0
                    
                    .ActMenge = 0#
                    .ActPreis = 0
                    If (ParenteralRezept = 24) Or (ParenteralRezept = 25) Then
                        .ActPreis = TeilMenge * 9.52
                        If (TeilMenge > 30) Then
                            .ActPreis = .ActPreis + (TeilMenge - 30) * 2.6
                            TeilMenge = 30
                        End If
                        If (TeilMenge > 15) Then
                            .ActPreis = .ActPreis + (TeilMenge - 15) * 3.7
                            TeilMenge = 15
                        End If
                        .ActPreis = .ActPreis + TeilMenge * IIf(ParenteralRezept = 24, 8.56, 9.52)
                    ElseIf (ParenteralRezept = 26) Or (ParenteralRezept = 27) Then
                        .ActPreis = TeilMenge * 5.8
                        .ActPreis = .ActPreis * 2#
                    Else
                        Dim AEKproML#
                        With frmTaxieren.flxTaxieren
                            For i% = 1 To (.Rows - 1)
                                ActFlag% = Val(.TextMatrix(i%, 6))
                                If (ActFlag% = MAG_SPEZIALITAET) Then
                                    AEKproML = .TextMatrix(i%, 7) / .TextMatrix(i%, 8)
                                    Exit For
                                End If
                            Next
                        End With
    '                        MsgBox (CStr(AEKproML))
                        
                        If (ParenteralRezept = 28) Then
                            .ActPreis = AEKproML * TeilMenge * 0.9
                            If (.ActPreis > 80) Then
                                TeilMenge = TeilMenge - (80 / (AEKproML * 0.9))
                                .ActPreis = 80 + (AEKproML * TeilMenge * 0.03)
                            End If
                        ElseIf (ParenteralRezept = 29) Then
                            If (TeilMenge > 16.495) Then
                                .ActPreis = AEKproML * (TeilMenge - 16.495) * 0.084
                                TeilMenge = 16.495
                            End If
                            .ActPreis = .ActPreis + TeilMenge * 4.85
                        ElseIf (ParenteralRezept = 30) Then
                            .ActPreis = AEKproML * TeilMenge * 0.9
                            If (.ActPreis > 100) Then
                                TeilMenge = TeilMenge - (100 / (AEKproML * 0.9))
                                .ActPreis = 100 + (AEKproML * TeilMenge * 0.03)
                            End If
                        End If
                    End If
                    
                    .flag = MAG_PREISEINGABE
                End With
                Call frmTaxieren.ZeigeTaxierZeile(.row)
            End If
        
            If (ParenteralRezept = 24) Or (ParenteralRezept = 26) Or (ParenteralRezept = 28) Or (ParenteralRezept = 30) Or (ParenteralRezept = 31) Then
                .AddItem " "
                .row = .Rows - 1
                row% = .row
                
                With TaxierRec
                    .pzn = Space$(Len(.pzn))
                    .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
                    .menge = Space$(Len(.menge))
                    .Meh = Space$(Len(.Meh))
                    .kp = 0
                    .GStufe = 0
                    
                    .ActMenge = 0#
                    .ActPreis = 8.35
                    
                    .flag = MAG_PREISEINGABE
                End With
                Call frmTaxieren.ZeigeTaxierZeile(.row)
            End If
                
            .AddItem " "
            .row = .Rows - 1
            row% = .row
            
            With TaxierRec
                .pzn = "02567001"   ' Space$(Len(.pzn))
                .kurz = Left$("BTM-Geb¸hr " + Space$(Len(.kurz)), Len(.kurz))
                .menge = Space$(Len(.menge))
                .Meh = Space$(Len(.Meh))
                .kp = 0
                .GStufe = 0
                
                .ActMenge = 1#
                .ActPreis = 3.58
                .kp = .ActPreis
                
                .flag = MAG_PREISEINGABE
            End With
            Call frmTaxieren.ZeigeTaxierZeile(.row)
        End If
    End If
End With
    
Call DefErrPop
End Sub

Sub TaxSumme()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaxSumme")
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
Dim i%, ActFlag%, AktMwSt%, FixAufschlagDa%, row%, aRow%, ArbeitDa%
Dim ActPreis#, ActMenge#, FaktorPreis#, FaktorPreisZuz#, SumGef‰ssPreis#, SumSpezPreis#
Dim h$, h2$
  
SumPreis# = 0#
TeilPreis# = 0#
ProzPreis# = 0#
SumPreisZuz# = 0#
TeilPreisZuz# = 0#
ProzPreisZuz# = 0#
FaktorPreis# = 0#
FaktorPreisZuz# = 0#
TeilMenge = 0#
UnverarbeiteteAbgabe = 0

SumGef‰ssPreis# = 0#
SumSpezPreis# = 0#

AktMwSt% = para.Mwst(2)

With frmTaxieren.flxTaxieren
'    If (ParenteralRezept > 15) Then
'        TeilMenge = 0
'        For i% = 1 To (.Rows - 1)
'            ActMenge# = iCDbl(.TextMatrix(i%, 1))
'            ActFlag% = Val(.TextMatrix(i%, 6))
'
'            If (ActFlag% < MAG_NN) Then
'                If (ActFlag% <> MAG_GEFAESS) And (ActFlag% <> MAG_ARBEIT) And (ActFlag% <> MAG_SONSTIGES) Then
'                    If ((ActFlag% <> MAG_SPEZIALITAET) And (ActFlag% <> MAG_ANTEILIG)) Or (.TextMatrix(i%, 2) <> "ST") Then
'                        TeilMenge# = TeilMenge# + ActMenge#
'                    End If
'                End If
'            End If
'        Next
'        If (TeilMenge > 0) Then
'            FixAufschlagDa = 0
'            For i% = 1 To (.Rows - 1)
'                If (InStr(UCase(.TextMatrix(i, 3)), "FIXZUSCHLAG") > 0) Then
'                    FixAufschlagDa = True
'                    Exit For
'                End If
'            Next
'
'            If (FixAufschlagDa = 0) Then
'                .AddItem " "
'                .row = .Rows - 1
'                row% = .row
'
'                With TaxierRec
'                    .pzn = Space$(Len(.pzn))
'                    .kurz = Left$("Fixzuschlag" + Space$(Len(.kurz)), Len(.kurz))
'                    .menge = Space$(Len(.menge))
'                    .meh = Space$(Len(.meh))
'                    .kp = 0
'                    .Gstufe = 0
'
'                    .ActMenge = 0#
'                    .ActPreis = 0
'                    If (ParenteralRezept = 16) Or (ParenteralRezept = 17) Then
'                        If (TeilMenge > 30) Then
'                            .ActPreis = .ActPreis + (TeilMenge - 30) * 2.6
'                            TeilMenge = 30
'                        End If
'                        If (TeilMenge > 15) Then
'                            .ActPreis = .ActPreis + (TeilMenge - 15) * 3.7
'                            TeilMenge = 15
'                        End If
'                        .ActPreis = .ActPreis + TeilMenge * IIf(ParenteralRezept = 16, 8.56, 9.52)
'                    Else
'                        Dim AEKproML#
'                        With frmTaxieren.flxTaxieren
'                            For i% = 1 To (.Rows - 1)
'                                ActFlag% = Val(.TextMatrix(i%, 6))
'                                If (ActFlag% = MAG_SPEZIALITAET) Then
'                                    AEKproML = .TextMatrix(i%, 7) / .TextMatrix(i%, 8)
'                                    Exit For
'                                End If
'                            Next
'                        End With
''                        MsgBox (CStr(AEKproML))
'
'                        If (ParenteralRezept = 18) Then
'                            .ActPreis = AEKproML * TeilMenge * 0.9
'                            If (.ActPreis > 80) Then
'                                TeilMenge = TeilMenge - (80 / (AEKproML * 0.9))
'                                .ActPreis = 80 + (AEKproML * TeilMenge * 0.03)
'                            End If
'                        ElseIf (ParenteralRezept = 19) Then
'                            If (TeilMenge > 16.495) Then
'                                .ActPreis = AEKproML * (TeilMenge - 16.495) * 0.084
'                                TeilMenge = 16.495
'                            End If
'                            .ActPreis = .ActPreis + TeilMenge * 4.85
'                        End If
'                    End If
'
'                    .flag = MAG_PREISEINGABE
'                End With
'                Call frmTaxieren.ZeigeTaxierZeile(.row)
'
'                If (ParenteralRezept = 16) Or (ParenteralRezept = 18) Then
'                    .AddItem " "
'                    .row = .Rows - 1
'                    row% = .row
'
'                    With TaxierRec
'                        .pzn = Space$(Len(.pzn))
'                        .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
'                        .menge = Space$(Len(.menge))
'                        .meh = Space$(Len(.meh))
'                        .kp = 0
'                        .Gstufe = 0
'
'                        .ActMenge = 0#
'                        .ActPreis = 8.35
'
'                        .flag = MAG_PREISEINGABE
'                    End With
'                    Call frmTaxieren.ZeigeTaxierZeile(.row)
'                End If
'            End If
'        End If
'    End If
    
    If (ParenteralRezept > 15) Then
    Else
        ArbeitDa = 0
        For i% = 1 To (.Rows - 1)
            ActFlag% = Val(.TextMatrix(i%, 6))
            If (ActFlag = MAG_ARBEIT) Then
                If (Left$(.TextMatrix(i, 3), 5) <> "UNVER") Then
                    ArbeitDa = i    'True
                    Exit For
                End If
            End If
        Next
        
        If (ArbeitDa) Then
    '        aRow = .row
            FixAufschlagDa = 0
            For i% = 1 To (.Rows - 1)
                If (InStr(UCase(.TextMatrix(i, 3)), "FIX-AUFSCHLAG") > 0) Then
                    FixAufschlagDa = True
                    Exit For
                End If
            Next
            
            If (ParenteralRezept >= 0) And (ParenteralRezept <= 15) Then
            Else
                If (FixAufschlagDa = 0) Then
                    .AddItem " ", ArbeitDa + 1 '.Rows
                    .row = ArbeitDa + 1 '.Rows - 1
                    row% = .row
                    
                    With TaxierRec
                        .pzn = Space$(Len(.pzn))
                        .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
                        .menge = Space$(Len(.menge))
                        .Meh = Space$(Len(.Meh))
                        .kp = 0
                        .GStufe = 0
                        
                        .ActMenge = 0#
                        .ActPreis = 8.35
                        
                        .flag = MAG_PREISEINGABE
                    End With
                    Call frmTaxieren.ZeigeTaxierZeile(.row)
                
                End If
            End If
    '        .row = aRow
        End If
    End If
    
    TeilMenge = 0#
    For i% = 1 To (.Rows - 1)
        ActPreis# = iCDbl(.TextMatrix(i%, 0))
        ActMenge# = iCDbl(.TextMatrix(i%, 1))
        ActFlag% = Val(.TextMatrix(i%, 6))
        
        If (ActFlag% < MAG_NN) Then
            TeilPreis# = TeilPreis# + ActPreis#
            
            If (ParenteralRezept >= 0) Then
'                If (ActFlag% = MAG_GEFAESS) Then
'                    SumGef‰ssPreis# = SumGef‰ssPreis# + ActPreis#
'                ElseIf (ActFlag% = MAG_SPEZIALITAET) Then
'                    SumSpezPreis# = SumSpezPreis# + ActPreis#
'                End If
            Else
                If (ActFlag% <> MAG_SONSTIGES) Then
                    ProzPreis# = ProzPreis# + ActPreis# * (AktMwSt% / 100#)
                End If
            
                If (ActFlag% <> MAG_GEFAESS) And (ActFlag% <> MAG_ARBEIT) And (ActFlag% <> MAG_SONSTIGES) Then
'                    If ((ActFlag% <> MAG_SPEZIALITAET) And (ActFlag% <> MAG_ANTEILIG)) Or (.TextMatrix(i%, 2) <> "ST") Then
                    If (.TextMatrix(i%, 2) = "ST") Then 'Or ((ActFlag% = MAG_SPEZIALITAET) Or (ActFlag% = MAG_ANTEILIG)) Then
                    Else
                        TeilMenge# = TeilMenge# + ActMenge#
                    End If
                End If
            End If
        
            If (IstFiveRxPzn(.TextMatrix(i%, 4))) Then
            Else
                TeilPreisZuz# = TeilPreisZuz# + ActPreis#
                
                If (ParenteralRezept >= 0) Then
                Else
                    If (ActFlag% <> MAG_SONSTIGES) Then
                        ProzPreisZuz# = ProzPreisZuz# + ActPreis# * (AktMwSt% / 100#)
                    End If
                End If
            End If
        
        End If
        
        If (Left$(.TextMatrix(i, 3), 5) = "UNVER") Then
            UnverarbeiteteAbgabe = i
        End If
        
    Next i%
'    SumPreis# = TeilPreis# + ProzPreis#
End With

With frmTaxieren.flxTaxSumme
    .Redraw = False
    .Rows = 0
    
    If (ParenteralRezept >= 0) Then
        h$ = Format(TeilPreis#, "0.00")
        .AddItem h$
        
'        If (SumSpezPreis# > 0) Then
'            ActPreis = SumSpezPreis# * 0.03
'            h$ = Format(ActPreis#, "0.00")
'            .AddItem h$ + vbTab + " " + vbTab + "+Aufschlag 3% FAM"
'            TeilPreis# = TeilPreis# + ActPreis#
'        End If
'        If (SumGef‰ssPreis# > 0) Then
'            ActPreis = SumGef‰ssPreis# * 0.15
'            h$ = Format(ActPreis#, "0.00")
'            .AddItem h$ + vbTab + " " + vbTab + "+Aufschlag 15% Gef‰ﬂ"
'            TeilPreis# = TeilPreis# + ActPreis#
'        End If
        
        Dim TeilPreis2#
        TeilPreis2 = TeilPreis
        With frmTaxieren.flxTaxieren
            For i = 1 To (.Rows - 1)
                If (InStr(UCase(.TextMatrix(i, 3)), UCase("Honorierung des Sichtbezuges")) > 0) Then
                    TeilPreis2 = TeilPreis2 - iCDbl(.TextMatrix(i%, 0))
                    Exit For
                End If
            Next i
        End With
        ActPreis = TeilPreis2# * (AktMwSt% / 100#)
'        ActPreis = TeilPreis# * (AktMwSt% / 100#)
        
        h$ = Format(ActPreis, "0.00")
        .AddItem h$ + vbTab + " " + vbTab + "+" + Format(AktMwSt%, "0") + "%"
        TeilPreis# = TeilPreis# + ActPreis#
        
        SumPreis = TeilPreis
    
        ActPreis = TeilPreisZuz# * (AktMwSt% / 100#)
        TeilPreisZuz# = TeilPreisZuz# + ActPreis#
        
        SumPreisZuz = TeilPreisZuz
    Else
        h$ = Format(TeilPreis#, "0.00")
        h$ = h$ + vbTab + Format(TeilMenge#, "0")
        .AddItem h$
        
        h$ = Format(ProzPreis#, "0.00")
        .AddItem h$ + vbTab + " " + vbTab + "+" + Format(AktMwSt%, "0") + "%"
        
        If (RezepturMitFaktor%) Then
            FaktorPreis# = (TeilPreis# + ProzPreis#) * (VmRabattFaktor# - 1#)
            h$ = Format(FaktorPreis#, "0.00")
            h2$ = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")
            If (Right$(h2$, 2) = "00") Then h2$ = Left$(h2$, Len(h2$) - 3)
            .AddItem h$ + vbTab + " " + vbTab + "+" + h2$ + "%"
            
            FaktorPreisZuz# = (TeilPreisZuz# + ProzPreisZuz#) * (VmRabattFaktor# - 1#)
        End If
        
        SumPreis# = TeilPreis# + ProzPreis# + FaktorPreis#
        SumPreisZuz# = TeilPreisZuz# + ProzPreisZuz# + FaktorPreisZuz#
    End If
    
    SumPreis# = (SumPreis# * 100#)
    SumPreis# = CLng(SumPreis#)
    SumPreis# = SumPreis# / 100#

    SumPreisZuz# = (SumPreisZuz# * 100#)
    SumPreisZuz# = CLng(SumPreisZuz#)
    SumPreisZuz# = SumPreisZuz# / 100#

'    h$ = Format(TeilPreis# + ProzPreis# + FaktorPreis#, "0.00")
    h$ = Format(SumPreis#, "0.00")
    .AddItem h$ + vbTab + vbTab + "Summe"
    
    .FillStyle = flexFillRepeat
    .row = .Rows - 1
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
    If (para.Newline = 0) Then
        .CellForeColor = vbBlack
        .CellBackColor = vbWhite
    End If
    .CellFontBold = True
    .FillStyle = flexFillSingle
        
    .Height = .RowHeight(0) * .Rows
    If (para.Newline = 0) Then
        .Height = .Height + 90
    End If
        
    .Redraw = True
End With
    
Call DefErrPop
End Sub

Function GPMenge#(sEinheit$, sMenge$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GPMenge!")
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
Dim mal%
Dim ret#

ret# = 1
Do
    mal% = InStr(sMenge$, "X")
    If (mal% > 0) Then
        If (Val(Left$(sMenge$, mal% - 1)) <> 0) Then
            ret = ret * Val(Left$(sMenge$, mal% - 1))
        End If
        sMenge$ = Mid$(sMenge$, mal% + 1)
    ElseIf (xVal(sMenge$) <> 0) Then
        ret = ret * xVal(sMenge$)
'        sMenge = ""
    End If
Loop Until (mal% = 0)
If (Trim(sEinheit) = "KG") Then
    ret = ret * 1000#
    sEinheit = "G "
ElseIf (Trim(sEinheit) = "L") Then
    ret = ret * 1000#
    sEinheit = "ML "
End If

GPMenge = ret

Call DefErrPop
End Function

Function iCDbl#(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iCDbl#")
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
Dim ret#

ret# = 0#
If (Trim(s$) <> "") Then ret# = CDbl(s$)

iCDbl# = ret#

Call DefErrPop
End Function

Function CheckUmlaute%(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckUmlaute%")
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
Dim i%, ret%, ind2%
Dim sSteuerFalsch$, sSteuerOk$

sSteuerFalsch = "ØÑôö"
sSteuerOk = "ﬂƒ÷‹"
    
ret = 0

For i = 1 To 4
    Do
        ind2 = InStr(s, Mid(sSteuerFalsch, i, 1))
        If (ind2 > 0) Then
            ret = True
            Mid(s, ind2, 1) = Mid(sSteuerOk, i, 1)
        Else
            Exit Do
        End If
    Loop
Next i

CheckUmlaute = ret

Call DefErrPop
End Function

Sub CheckBtmGebuehr()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBtmGebuehr")
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
Dim i%, BtmDa%
  
BtmDa = 0
With frmTaxieren.flxTaxieren
    For i% = 1 To (.Rows - 1)
        If (InStr(UCase(.TextMatrix(i, 3)), "BTM-GEB‹HR") > 0) Then
            BtmDa = True
            .RemoveItem (i)
            Exit For
        End If
    Next
    
    If (BtmDa = 0) Then
        Dim OrgRow%
        
        OrgRow = .row
        If (.row = .Rows - 1) Then
            .AddItem " "
        End If
        .row = .Rows - 1
      
        With TaxierRec
            .pzn = "02567001"   ' Space$(Len(.pzn))
            .kurz = Left$("BTM-Geb¸hr " + Space$(Len(.kurz)), Len(.kurz))
            .menge = Space$(Len(.menge))
            .Meh = Space$(Len(.Meh))
            .kp = 0
            .GStufe = 0
            
            .ActMenge = 1#
            .ActPreis = 3.58
            .kp = .ActPreis
            
            .flag = MAG_PREISEINGABE
        End With
        Call frmTaxieren.ZeigeTaxierZeile(.row)
    
        .row = OrgRow
    End If
End With
    
Call DefErrPop
End Sub

Public Sub HoleFiveRxPzns()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleFiveRxPzns")
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
Dim l&
Dim h$, Key$
  
MaxFiveRxPzns = -1

For i% = 1 To 50
    h$ = Space$(50)
    Key$ = "Pzn" + Format(i%, "0")
    l& = GetPrivateProfileString("Sonderfaelle", Key$, h$, h$, 51, CurDir + "\FiveRxPzn.ini")
    h$ = Left$(h$, l&)
    
    If (h = "") Then
        Exit For
    End If
    
    MaxFiveRxPzns = i - 1
    FiveRxPzns(i - 1) = h
Next i%

If (FiveRxPzns(0) = "") Then
    MaxFiveRxPzns = 3
    FiveRxPzns(0) = "09999637,Beschaffungskosten,2.98"
    FiveRxPzns(1) = "06461110,Botendienst,2.98"
    FiveRxPzns(2) = "02567001,BTM-Geb¸hr,4.26"
    FiveRxPzns(3) = "02567018,Noctu-Geb¸hr,2.50"
End If

Call DefErrPop
End Sub

Public Function IstFiveRxPzn%(sPzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstFiveRxPzn%")
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
Dim i%, ind%, ret%
Dim h$, h2$
  
ret = 0

For i = 0 To MaxFiveRxPzns
    h = FiveRxPzns(i)
    ind = InStr(h, ",")
    If (ind > 0) Then
        h2 = Trim(Left(h, ind - 1))
        If (Val(h2) = Val(sPzn)) Then
            ret = True
            Exit For
        End If
    End If
Next

IstFiveRxPzn = ret

Call DefErrPop
End Function



