Attribute VB_Name = "modInfo"
Option Explicit

Type FileArray
   kurz As String * 15
   Bez As String * 50
   Name As String * 50
   Typ As Integer
   Laenge As Integer
   FeldAnz As Integer
   fh As Integer
   AnzDS As Long
   AktDS As Long
   DS As String * 255
   FTab As String * 1
   Pointer As Integer
   Offset As Integer
   SortField As Integer
End Type

Type FieldArr
   kurz As String * 20
   Bez As String * 30
   Start As Integer
   Laenge As Integer
   Typ As String * 1
   Konv As Integer
   RelStart As Integer
   RelAnz As Integer
   ind As Integer
End Type

Public Const MAX_INFO_ZEILEN = 6
Public Const MIN_INFO_ZEILEN = 4

Const MAXINDEX = 100
Const gcsIniFile$ = "LISTEN.INI"

Public DateiTab(5) As FileArray
Public Feldtab(MAXINDEX, 5) As FieldArr

Public InfoLayoutInd%

Dim recTaxe As Recordset
Dim StammFix As String * 100

Dim DateienEingelesen%

Private Const DefErrModul = "winana.bas"

Sub EinlesenAlleDateiHeader()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenAlleDateiHeader")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ret%

ret% = EinlesenDateiHeader%(0, "TAXE")
ret% = EinlesenDateiHeader%(1, "AST")
ret% = EinlesenDateiHeader%(2, "ASS")
ret% = EinlesenDateiHeader%(3, "ASSLIF")
ret% = EinlesenDateiHeader%(4, "ASSVK")
ret% = EinlesenDateiHeader%(5, "ASSTARA")

For i% = 0 To 5
    Call EinlesenDateiFelder(i%)
Next i%

Call DefErrPop
End Sub

Function EinlesenDateiHeader%(DateiInd%, kurz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenDateiHeader%")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, iRet%, anz%, iAnzFiles%
Dim l&
Dim sDat$, zw$, s$()

If (kurz$ = "TAXE") Then
    DateiTab(DateiInd%).kurz = kurz$
    DateiTab(DateiInd%).Bez = "Arzneimitteltaxe"
    EinlesenDateiHeader% = True
    Call DefErrPop: Exit Function
End If

sDat$ = Space$(100)
l& = GetPrivateProfileString("DATEIEN", kurz$, "", sDat$, 101, "\user\" + gcsIniFile$)
sDat$ = Trim$(sDat$)
If (Len(sDat$) <= 1) Then
    EinlesenDateiHeader% = False
    Call DefErrPop: Exit Function
End If

sDat$ = Left$(sDat$, Len(sDat$) - 1)
ReDim s$(0)

anz% = 5
iRet% = ZerlegeIniEintrag%(s$(), anz%, sDat$)
If (iRet%) Then
    DateiTab(DateiInd%).kurz = kurz$
    If (UCase(kurz$) = "AST") Then s$(1) = "Lagerartikel"
    DateiTab(DateiInd%).Bez = s$(1)
    DateiTab(DateiInd%).Name = s$(2)
    j% = InStr(s$(3), "x")
    If (j% > 0) Then
        DateiTab(DateiInd%).Offset = Val(Mid$(s$(3), j% + 1))
        s$(3) = Left$(s$(3), j% - 1)
    End If
    DateiTab(DateiInd%).Typ = Val(s$(3))
    DateiTab(DateiInd%).Laenge = Val(s$(4))
    zw$ = RTrim$(s$(5))
    If (zw$ > "") Then       'Pointer auf Ursprungsdatei
        For i% = 0 To (DateiInd% - 1)
            If (RTrim$(DateiTab(i%).kurz) = zw$) Then
                'Index der Ursprungsdatei speichern
                DateiTab(DateiInd%).Pointer = i%
                Exit For
            End If
        Next i%
    ElseIf (DateiTab(DateiInd%).Typ = 4) Then   'Taxe
        DateiTab(DateiInd%).Name = TaxeLw$ + ":\TAXE\" + DateiTab(DateiInd%).Name
    End If
End If

EinlesenDateiHeader% = True

Call DefErrPop
End Function

Function ZerlegeIniEintrag%(s$(), anz%, sDat$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZerlegeIniEintrag%")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%

If (anz% > UBound(s$)) Then ReDim s$(anz%)
ZerlegeIniEintrag% = True
For i% = 1 To anz%
   j% = InStr(sDat$, ";")
   If (j% > 0) Then
      s$(i%) = Left$(sDat$, j% - 1)
      sDat$ = Mid$(sDat$, j% + 1)
   Else
      If (i% = anz%) Then
         s$(i%) = sDat$
      Else
         ZerlegeIniEintrag% = False
         Exit For
      End If
   End If
Next i%
Call DefErrPop
End Function

Function SucheInfoFeld$(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheInfoFeld$")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim trenn%, Such%, DateiInd%, f%
Dim kurz$, sInfoFeld$, Datei$

If (DateienEingelesen% = False) Then
    Call EinlesenAlleDateiHeader
    DateienEingelesen% = True
End If

sInfoFeld$ = ""

DateiInd% = -1: f% = -1
trenn% = InStr(s$, ".")
If (trenn% > 1) Then
    Datei$ = UCase$(Left$(s$, trenn% - 1))
    s$ = UCase$(RTrim$(Mid$(s$, trenn% + 1)))
    
    For Such% = 0 To UBound(DateiTab)
        kurz$ = UCase$(DateiTab(Such%).kurz)
        kurz$ = RTrim(kurz$)
        If (kurz$ = Datei$) Then
            DateiInd% = Such%
            Exit For
        End If
    Next Such%
    
    If (DateiInd% >= 0) Then
'        If (Datei$ = "TAXE") Then
'            sInfoFeld$ = UCase(Datei$ + "." + s$)
'            SucheInfoFeld$ = sInfoFeld$
'            Call DefErrPop: Exit Function
'        End If
        
        For Such% = 0 To DateiTab(DateiInd%).FeldAnz
            kurz$ = UCase$(Feldtab(Such%, Asc(DateiTab(DateiInd%).FTab)).kurz)
            kurz$ = RTrim(kurz$)
            If (kurz$ = s$) Then
                f% = Such%
                sInfoFeld$ = Str$(DateiInd% * 100 + (f%))
                Exit For
            End If
        Next Such%
    End If
End If

SucheInfoFeld$ = sInfoFeld$
  
  
'????  If f% >= 0 Then Inhalt$ = str$((Datei% + 1) * 100 + (f%))
'    IF Inhalt$ > "" OR (ITyp <> FELD AND ITyp <> VFELD AND ITyp <> MAKRO) TH
'      anz% = anz% + 1
'      ReDim Preserve c(anz%) As Ausdruck
'      c(anz%).typ = ITyp
'      c(anz%).Inhalt = Inhalt$
'   If Bereich% = Ausgabe Then
'        c(anz%).DruckLen = val(Flen$)
'        c(anz%).KopfZ = KopfZ$
'      End If
'      ok% = True
'   End If


Call DefErrPop
End Function

Sub EinlesenDateiFelder(iFileNo%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenDateiFelder")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim j%, iAnzFields%, iRet%, anz%
Dim l&
Dim s$(), sDat$

iAnzFields% = 0

ReDim s$(0)

If (iFileNo% = 0) Then
    iAnzFields% = TaxeDB.TableDefs("Taxe").Fields.Count
    For j% = 0 To (iAnzFields% - 1)
        sDat$ = RTrim(TaxeDB.TableDefs("Taxe").Fields(j%).Name)
        Call OemToChar(sDat$, sDat$)
        Feldtab(j%, iFileNo%).kurz = sDat$
        Feldtab(j%, iFileNo%).Bez = sDat$
'        c.AddItem FirstLettersUcase$(RTrim(TaxeDB.TableDefs("Taxe").Fields(j%).Name))
    Next j%
Else
    Do While True
        sDat$ = Space$(100)
        l& = GetPrivateProfileString(DateiTab(iFileNo%).kurz, Mid$(Str$(iAnzFields% + 1), 2), "", sDat$, 101, "\user\" + gcsIniFile$)
        sDat$ = Trim$(sDat$)
        If (Len(sDat$) <= 1) Then Exit Do
    
        sDat$ = Left$(sDat$, Len(sDat$) - 1)
        Call OemToChar(sDat$, sDat$)
        
        anz% = 5
        iRet% = ZerlegeIniEintrag%(s$(), anz%, sDat$)
        If (iRet% And Len(sDat$) > 0) Then
            Feldtab(iAnzFields%, iFileNo%).kurz = s$(1)
            Feldtab(iAnzFields%, iFileNo%).Bez = s$(2)
            Feldtab(iAnzFields%, iFileNo%).Start = Val(s$(3))
            Feldtab(iAnzFields%, iFileNo%).Laenge = Val(s$(4))
            Feldtab(iAnzFields%, iFileNo%).Typ = s$(5)
    '        jTmp% = 0: GoSub NextFEintrag
    '        If (jTmp% > 0) Then
    '            Feldtab(iAnzFields%, iFileNo%).Konv = val(Left$(sDat$, jTmp% - 1))
    '            GoSub NextFEintrag
    '        Else
                Feldtab(iAnzFields%, iFileNo%).Konv = Val(sDat$)
    '        End If
            Feldtab(iAnzFields%, iFileNo%).ind = -1
    '        If (jTmp% > 0) Then GoSub RelationEinlesen
        End If
        iAnzFields% = iAnzFields% + 1
    Loop
    'iRet = GetPrivProfileStr(DateiTab(iFileNo%).kurz, "SORT", "", s$, 512, USERPFAD$ + gcsIniFile$)
    'If (iRet <> 0) Then
    '  If (val(s$) <= iAnzFields%) Then DateiTab(iFileNo%).SortField = val(s$)
    'End If
    'GoSub IndexEinlesen
End If

DateiTab(iFileNo%).FeldAnz = iAnzFields%
DateiTab(iFileNo%).FTab = Chr$(iFileNo%)

Call DefErrPop
End Sub

Function FirstLettersUcase(ByVal s As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FirstLettersUcase")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  
Dim i As Integer
Dim c As String
  
s = UCase(Left(s, 1)) + LCase(Mid(s, 2))
For i = 2 To Len(s)
  c = UCase(Mid(s, i - 1, 1))
  If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜß", c) = 0 Then
    Mid(s, i, 1) = UCase(Mid(s, i, 1))
  End If
Next i
FirstLettersUcase = s

Call DefErrPop
End Function

Function EinlesenDateiDS(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenDateiDS")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, kurz$, SQLStr$

For i% = 0 To UBound(DateiTab)
    kurz$ = UCase$(DateiTab(i%).kurz)
    kurz$ = RTrim(kurz$)
    
    Select Case (kurz$)
        Case "TAXE"
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
            Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
        Case "ASS", "ASSLIF"
            FabsErrf% = IndexSearch(STATISTIK, 0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Seek #OpDateien(STATISTIK).Handle, FabsRecno& + 1
                Get #OpDateien(STATISTIK).Handle, , DateiTab(i%).DS
            Else
                DateiTab(i%).DS = String$(255, 64)
            End If
        Case "AST"
            FabsErrf% = IndexSearch(STAMM, 0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Seek #OpDateien(STAMM).Handle, FabsRecno& + 1
                Get #OpDateien(STAMM).Handle, , StammFix
                DateiTab(i%).DS = StammFix
            Else
                DateiTab(i%).DS = String$(255, 64)
            End If
    End Select
Next i%

Call DefErrPop
End Function

Function HoleTaxeInfoWert$(FeldName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleTaxeInfoWert$")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim h$
Dim fldRecordset As Field

If (recTaxe.EOF = True) Then
    HoleTaxeInfoWert$ = ""
    Call DefErrPop: Exit Function
End If

Set fldRecordset = recTaxe.Fields(FeldName$)
Select Case UCase(FeldName$)
    Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT"
        HoleTaxeInfoWert$ = Format(fldRecordset.Value / 100#, "0.00")
        If (HoleTaxeInfoWert$ = "0.00") Or (HoleTaxeInfoWert$ = "0,00") Then
            HoleTaxeInfoWert$ = "0"
        End If
        HoleTaxeInfoWert$ = " " + HoleTaxeInfoWert$
    Case Else
        HoleTaxeInfoWert$ = " " + LTrim$(fldRecordset.Value)
End Select

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%
Call DefErrPop
End Function

Function HoleInfoWert$(DateiInd%, FeldInd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleInfoWert$")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iStart%, iLaenge%, iKonv%
Dim iTyp$, h$

If (DateiInd% = 0) Then
    HoleInfoWert$ = HoleTaxeInfoWert$(RTrim(Feldtab(FeldInd%, DateiInd%).kurz))
Else
    iStart% = Feldtab(FeldInd%, DateiInd%).Start
    iLaenge% = Feldtab(FeldInd%, DateiInd%).Laenge
    iTyp$ = UCase(Feldtab(FeldInd%, DateiInd%).Typ)
    iKonv% = Feldtab(FeldInd%, DateiInd%).Konv
    
    If (Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%) = String$(iLaenge%, 64)) Then
        HoleInfoWert$ = ""
        Call DefErrPop: Exit Function
    End If
    
    Select Case iTyp$
        Case "T"
            HoleInfoWert$ = " " + LTrim$(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%))
        Case "N"
            Select Case iKonv%
                Case 1
                    'ASC
                    HoleInfoWert$ = Str$(Asc(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)))
                 Case 2
                    'CVD
                    HoleInfoWert$ = Format$(CVD#(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)), "0.00")
'                    HoleInfoWert$ = Str$(CVD#(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)))
                Case 3
                    'VAL
                    HoleInfoWert$ = Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)
                 Case 4
                    'CVS
                    HoleInfoWert$ = Format$(CVS!(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)), "0.00")
'                    HoleInfoWert$ = Str$(CVS!(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)))
                 Case 5
                    'CVI
                    HoleInfoWert$ = Str$(CVI%(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)))
                 Case 6
                    'CVI / 10
                    HoleInfoWert$ = Format$(CVI%(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)) / 10, "0.00")
'                    HoleInfoWert$ = Str$(CVI%(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%)) / 10)
            End Select
            HoleInfoWert$ = LTrim$(RTrim$(HoleInfoWert$))
            If (HoleInfoWert$ = "0.00") Or (HoleInfoWert$ = "0,00") Then
                HoleInfoWert$ = "0"
            End If
            HoleInfoWert$ = " " + HoleInfoWert$
        Case "D"
            Select Case iKonv%
                Case 10
                    'CVDatum$
                    h$ = CVDatum$(Mid$(DateiTab(DateiInd%).DS, iStart%, iLaenge%))
                    HoleInfoWert$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
            End Select
    End Select
End If

Call DefErrPop
End Function

'Function KonvertFeld$(s$, Art%, Typ As String, Konv As Integer)
'
'Select Case Typ
'Case "N", "n"
'  Art% = NUMERISCH
'  Select Case Konv
'  Case 1
'    s$ = LTrim$(str$(Asc(s$)))
'    Art% = N3V0
'  Case 2
'    s$ = LTrim$(str$(FNX#(CVD(s$))))
'    Art% = N8V2
'  Case 3, 19, 8
'    'VAL
'    If Konv = 8 Then
'      Art% = N8V2
'      If Len(s$) > 2 Then
'        s$ = Left$(s$, Len(s$) - 2) + "." + Right$(s$, 2)
'      Else
'        s$ = "." + Right$("00" + LTrim$(s$), 2)
'      End If
'    ElseIf InStr(s$, ".") <> 0 Then
'      Art% = N8V2
'    Else
'      Art% = N5V0
'    End If
'  Case 4
'    s$ = LTrim$(str$(FNX#(CVS(s$))))
'    Art% = N8V2
'  Case 5
'    s$ = LTrim$(str$(CVI(s$)))
'    Art% = N5V0
'  Case 6
'    s$ = LTrim$(str$(CVI(s$) / 10))
'    Art% = N8V2
'  Case 12       '1.19
'    s$ = Right$("0000" + Mid$(str$(CVI(s$)), 2), 4)
'    s$ = Left$(s$, 2) + ":" + Mid$(s$, 3, 2)
'    Art% = ALPHA
'  Case 13
'    If MwSt!(1) = 0 Then
'      f.tmp% = FreeFile
'      Open "pdatei.dat" For Random Access Read Shared As #f.tmp% Len = 30
'        FIELD #f.tmp%, 30 AS pda$
'      GET #f.tmp%, 2
'      For i% = 1 To 4
'        MwSt!(i%) = Asc(Mid$(pda$, 14 + i%, 1))
'      Next i%
'      Close #f.tmp%
'    End If
'    If val(s$) < 4 Then
'      s$ = LTrim$(str$(MwSt!(val(s$))))
'    Else
'      s$ = LTrim$(str$(MwSt!(2)))
'    End If
'    Art% = N3V0
'  Case Else
'    KonvTyp$ = Mid$(str$(Konv), 2, 1)
'    If KonvTyp$ = "7" Or KonvTyp$ = "9" Then
'      wieviel% = val(Mid$(str$(Konv), 3))
'      If wieviel% Mod 2 = 1 Then wieviel% = wieviel% + 1
'      norm$ = String$(wieviel%, " ")    'MachNormal
'      If s$ <> String$(Len(s$), Asc("*")) Then  'bei FB "***"
'        Call Bcd2ascii(norm$, s$, wieviel%)
'      End If
'      If wieviel% > val(Mid$(str$(Konv), 3)) Then
'        wieviel% = val(Mid$(str$(Konv), 3))
'      End If
'        If wieviel% = 8 Or wieviel% = 6 Then
'          If KonvTyp$ = "7" Then
'            s$ = LTrim$(str$(val(norm$) / 100))
'          Else
'            s$ = Left$(norm$, wieviel%)
'          End If
'          Art% = N8V2
'        Else
'          If KonvTyp$ = "7" Then
'            s$ = LTrim$(str$(val(norm$)))
'          Else
'            s$ = Left$(norm$, wieviel%)
'          End If
'          Art% = N5V0
'        End If
'    End If
'  End Select
'Case "D", "d"
'  Art% = datum
'  Select Case Konv
'  Case 10
'    iDat% = CVDate%(s$)
'  Case 11
'    iDat% = CVI(s$)
'  Case 13
'    iDat% = iDate%(s$)
'  End Select
'  s$ = sDate$(iDat%)
'  If Len(s$) < 6 Then
'    s$ = Space$(8)
'  Else
'    s$ = DateLong$(s$)
'  End If
'Case Else
'  Art% = ALPHA
'End Select
'KonvertFeld$ = s$
'End Function


