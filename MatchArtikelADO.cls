VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMatchArtikelADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const ANZMATCHKONVERTIERUNGEN = 24

Const LANGERNAME = 0
Const AEPundAVP = 1
Const PZNundAVP = 2
Const FESTBETRAG = 3

Const INI_SECTION = "Auswahl-Artikel"
Const ARBEIT_SECTION = "Arbeitsbereich Auswahl-Artikel"
Const INFO_SECTION = "Infobereich Auswahl-Artikel"

Const EINGABE_NAME = "&Name/Pzn"

Dim MatchConn As ADODB.Connection
Dim MatchDB$
Dim SqlServerDB%

Dim MatchTaxeRec As New ADODB.Recordset
Dim MatchTaxeRec2 As New ADODB.Recordset

Dim TaxeAuswahlModus%

Dim MatchKlein$, MatchGross$, MatchGueltig$
Dim MatchPhonAlt$(ANZMATCHKONVERTIERUNGEN)
Dim MatchPhonNeu$(ANZMATCHKONVERTIERUNGEN)

Dim FabsErrf%
Dim FabsRecno&

'''''''
Dim iMatchModus%
Dim iMatchTyp%
Dim iMatchListeTyp%
Dim iMatchAnzeigeTyp$()

Dim iMatchAutoRet%
Dim iMatchNurEiner%

Dim iOrgSuch$

Dim iMatchcodePzn$
Dim iMatchcodeTxt$
Dim iMatchcodeErg$
''''
Dim MusterSucheBreak As Boolean
Dim InMusterSuche As Boolean
Dim InMatchcodeBlättern As Boolean

Dim Muster As String
Dim Packung As Integer
Dim SuchNGr As String
Dim herst As String
Dim MusterSuche As Boolean
Dim MusterAnf As String
Dim MusterName As String

Dim iWarteMaske%
Dim iAbbruch%

Dim SqlTaxeSuch$

Dim ErsterSortierName$
Dim SelectStr$, SelectStr2$, CountStr$

Private Const DefErrModul = "MATCHARTIKELADO.CLS"

Sub InitMatch()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitMatch")
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
ReDim iMatchAnzeigeTyp$(3)
iMatchAnzeigeTyp$(0) = "Taxe"
iMatchAnzeigeTyp$(1) = "Taxe phonetisch"
iMatchAnzeigeTyp$(2) = "Lagerartikel"
iMatchAnzeigeTyp$(3) = "Lagerartikel inkl. Merkzettel"

MatchKlein = "abcdefghijklmnopqrstuvwxyzäöüß,.-;:+*/ "
MatchGross = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜS........."
MatchGueltig = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ0123456789."
 
MatchPhonAlt$(0) = "J": MatchPhonNeu$(0) = "I"
MatchPhonAlt$(1) = "PH": MatchPhonNeu$(1) = "F"
MatchPhonAlt$(2) = "Y": MatchPhonNeu$(2) = "UE"
MatchPhonAlt$(3) = "Ä": MatchPhonNeu$(3) = "AE"
MatchPhonAlt$(4) = "Ö": MatchPhonNeu$(4) = "OE"
MatchPhonAlt$(5) = "Ü": MatchPhonNeu$(5) = "UE"
MatchPhonAlt$(6) = "SCH": MatchPhonNeu$(6) = "S"
MatchPhonAlt$(7) = "CH": MatchPhonNeu$(7) = "K"
MatchPhonAlt$(8) = "H": MatchPhonNeu$(8) = ""
MatchPhonAlt$(9) = "OE": MatchPhonNeu$(9) = "O"
MatchPhonAlt$(10) = "UE": MatchPhonNeu$(10) = "I"
MatchPhonAlt$(11) = "AI": MatchPhonNeu$(11) = "EI"
MatchPhonAlt$(12) = "AI": MatchPhonNeu$(12) = "EI"
MatchPhonAlt$(13) = "AE": MatchPhonNeu$(13) = "E"
MatchPhonAlt$(14) = "IE": MatchPhonNeu$(14) = "I"
MatchPhonAlt$(15) = "P": MatchPhonNeu$(15) = "B"
MatchPhonAlt$(16) = "T": MatchPhonNeu$(16) = "D"
MatchPhonAlt$(17) = "F": MatchPhonNeu$(17) = "V"
MatchPhonAlt$(18) = "V": MatchPhonNeu$(18) = "W"
MatchPhonAlt$(19) = "Z": MatchPhonNeu$(19) = "C"
MatchPhonAlt$(20) = "G": MatchPhonNeu$(20) = "K"
MatchPhonAlt$(21) = "C": MatchPhonNeu$(21) = "K"
MatchPhonAlt$(22) = "M": MatchPhonNeu$(22) = "N"
MatchPhonAlt$(23) = "O": MatchPhonNeu$(23) = "U"

'MatchModus% = TAXE_ALPHA
TaxeAuswahlModus% = AEPundAVP

Call clsError.DefErrPop
End Sub

Function MakeMatchCode$(Name$, Optional mModus% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeMatchCode$")
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
Dim ch$, chAlt$, h$, h2$, zw$, i%, l%, p%, ind%, Vnum$

If (mModus% = -1) Then mModus% = iMatchModus%

h$ = RTrim$(Name$)
l% = Len(h$)

If (mModus% = LAGER_MATCH) Or (mModus% = MERKZETTEL_MATCH) Then
'    ind% = InStr(Name$, vbTab)
'    If (ind% > 0) Then
'        Vnum$ = Mid$(Name$, ind% + 1)
'        Name$ = Left$(Name$, ind% - 1)
'    End If
    
    h$ = Name$
    l% = Len(h$)
    
    h2$ = UCase(h$)
    If (Para1.Land = "D") Then
'        If (l% > 28) Then l% = 28
        zw$ = ""
        For i% = 1 To l%
            ch$ = Mid$(h2$, i%, 1)
            If (ch$ = "ß") Then
                ch$ = "S"
            End If
            If (ch$ >= "A" And ch$ <= "Z") Or (ch$ >= "0" And ch$ <= "9") Or InStr("ÄÖÜß", ch$) > 0 Then
                zw$ = zw$ + ch$
            End If
        Next i%
        h2$ = zw$ + Space$(l% - Len(zw$))
'        If (Len(h$) > 28) Then h2$ = h2$ + Mid$(h$, 29)
        
'        Vnum$ = MakeVnum$(Vnum$)
'        MakeMatchCode$ = Left$(h2$ + Space$(10), 10) + Vnum$ + Mid$(h2$, 11)
        MakeMatchCode$ = h2$
    Else
'        Call OemToChar(h$, h$)
        h2$ = UCase(h$)
        For i% = 1 To l%
            ch$ = Mid$(h2$, i%, 1)
            If (ch$ = "ß") Then
                ch$ = "S"
                Mid$(h2$, i%, 1) = ch$
            End If
            If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜS0123456789", ch$) = 0) Then
                Mid$(h2$, i%, 1) = " "
            End If
        Next i%
'        Call CharToOem(h2$, h2$)
        MakeMatchCode$ = h2$
    End If
    
    Call clsError.DefErrPop: Exit Function
End If

For i% = 1 To l%
    ind% = InStr(MatchKlein$, Mid$(h$, i%, 1))
    If (ind% > 0) Then Mid$(h$, i%, 1) = Mid$(MatchGross$, ind%, 1)
Next i%

h2$ = ""
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    ind% = InStr(MatchGueltig$, ch$)
    If (ind% > 0) Then h2$ = h2$ + ch$
Next i%

h$ = h2$

For i% = 0 To (ANZMATCHKONVERTIERUNGEN - 1)
    p% = 2
    Do
        ind% = InStr(p%, h$, MatchPhonAlt(i%))
        If (ind% = 0) Then Exit Do
        If (Mid$(h$, ind% - 1, 1) <> ".") Then
            zw$ = Mid$(h$, ind% + Len(MatchPhonAlt(i%)))
            h$ = Left$(h$, ind% - 1) + MatchPhonNeu(i%) + zw$
        Else
            p% = p% + 1
        End If
    Loop
Next i%

l% = Len(h$)
h2$ = ""
chAlt$ = Chr$(1)
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    If (ch$ <> chAlt$) Then h2$ = h2$ + ch$
    chAlt$ = ch$
Next i%

h$ = Left$(h2$ + Space$(10), 10)
If (mModus% = LAGER_MATCH) Then h$ = Left$(h$, 6)

MakeMatchCode$ = RTrim$(h$)

Call clsError.DefErrPop
End Function

Function MakeVnum$(Vnum$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeVnum$")
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
Dim iVnum%, iVnum1%, iVnum2%, iVnum3%
Dim dVnum#

dVnum# = Val(Right$(Vnum$, 7)): iVnum% = Int(dVnum# / 128 / 128)
dVnum# = dVnum# - 128# * 128# * CDbl(iVnum%): iVnum1% = (iVnum% Mod 128)
iVnum2% = Int(dVnum# / 128): iVnum3% = dVnum# - 128 * iVnum2%
MakeVnum$ = Chr$(127 + iVnum1%) + Chr$(127 + iVnum2%) + Chr$(127 + iVnum3%)

Call clsError.DefErrPop
End Function
                      
Sub Umspeichern(buf$, pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Umspeichern")
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
Dim ind%, ind2%, i%, z%, z1%, z2%, dAnzFields%, IstMerkzettelMdb%
Dim zu!
Dim h$, h2$, char$, SQLStr$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

'Select Case iMatchModus%
'    Case TAXE_ALPHA, TAXE_MATCH
        
        Ausgabe(pos%).pzn = MatchTaxeRec!pzn
        
        h$ = " "
        SQLStr$ = "SELECT * FROM ArtikelInfo WHERE PZN = " + clsOpTool.PznString(MatchTaxeRec!pzn)
        MatchTaxeRec2.Open SQLStr$, ArtikelDB1.ActiveConn
        If (MatchTaxeRec2.EOF = False) Then
            Ausgabe(pos%).ZusatzTextKz = 2
            h$ = "+"
        End If
        MatchTaxeRec2.Close
        
        Ausgabe(pos%).LagerKz = 0
        If (MatchDB = "Artikel") Then
            If (MatchTaxeRec!LagerKz = 0) Then
                Ausgabe(pos%).BesorgtKz = 2
                If (Ausgabe(pos%).ZusatzTextKz = 2) Then
                    h$ = "±"
                Else
                    h$ = "-"
                End If
            ElseIf (UCase(MatchTaxeRec!halt) <> "S") Or (MatchTaxeRec!poslag > 0) Then
                Ausgabe(pos%).LagerKz = 2
            End If
        Else
            SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + clsOpTool.PznString(MatchTaxeRec!pzn)
            MatchTaxeRec2.Open SQLStr$, ArtikelDB1.ActiveConn
            If (MatchTaxeRec2.EOF = False) Then
                If (MatchTaxeRec2!LagerKz = 0) Then
                    Ausgabe(pos%).BesorgtKz = 2
                    If (Ausgabe(pos%).ZusatzTextKz = 2) Then
                        h$ = "±"
                    Else
                        h$ = "-"
                    End If
                ElseIf (UCase(MatchTaxeRec2!halt) <> "S") Or (MatchTaxeRec2!poslag > 0) Then
                    Ausgabe(pos%).LagerKz = 2
                End If
            End If
            MatchTaxeRec2.Close
        End If
        
'        h$ = " "
'        FabsErrf% = Zus1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
'        If (FabsErrf% = 0) Then
'            Ausgabe(pos%).ZusatzTextKz = 2
'            h$ = "+"
'        End If
'
'        On Error Resume Next
'        dAnzFields% = MerkzettelDB.TableDefs("Merkzettel").Fields.Count
'        IstMerkzettelMdb% = (Err = 0)
'        On Error GoTo DefErr
'
'        If (IstMerkzettelMdb%) Then
'            h2$ = Format(Ausgabe(pos%).pzn, "0000000")
'            SQLStr$ = "SELECT * FROM Merkzettel WHERE Pzn='" + clsOpTool.SqlPzn(h2$, 0) + "'"
'            Set MerkzettelRec = MerkzettelDB.OpenRecordset(SQLStr$)
'            If (MerkzettelRec.RecordCount > 0) Then
'                Ausgabe(pos%).BesorgtKz = 2
'                If (Ausgabe(pos%).ZusatzTextKz = 2) Then
'                    h$ = "±"
'                Else
'                    h$ = "-"
'                End If
'            End If
'        Else
'            FabsErrf% = Besorgt1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
'            If (FabsErrf% = 0) Then
'                Ausgabe(pos%).BesorgtKz = 2
'                If (Ausgabe(pos%).ZusatzTextKz = 2) Then
'                    h$ = "±"
'                Else
'                    h$ = "-"
'                End If
'            End If
'        End If
        
        Ausgabe(pos%).Name = h$ + vbTab + MachTaxeAuswahlName$(MatchTaxeRec)
        Ausgabe(pos%).Verweis = Len(RTrim(Ausgabe(pos%).Name))
        
'        Ausgabe(pos%).LagerKz = 0
'        FabsErrf% = Ass1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "0000000"), FabsRecno&)
'        If (FabsErrf% = 0) Then
'            Ass1.GetRecord (FabsRecno& + 1)
'            If (UCase(Ass1.halt) <> "S") Or (Ass1.poslag > 0) Then
'                Ausgabe(pos%).LagerKz = 2
'            End If
'        End If
        
        Ausgabe(pos%).key = MatchTaxeRec!SortierName
        Ausgabe(pos%).Bookmark = MatchTaxeRec.Bookmark
            
Call clsError.DefErrPop
End Sub

Function MachTaxeAuswahlName$(tRec As ADODB.Recordset)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachTaxeAuswahlName$")
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
Dim i%, iSieheKz%
Dim h$

iSieheKz = 0
If (iMatchModus <> LAGER_MATCH) And (iMatchModus <> MERKZETTEL_MATCH) Then
    iSieheKz = tRec!SieheKz
End If

If (iSieheKz) Then
    h$ = tRec!Name
    For i% = 1 To frmMatchcode!flxarbeit(0).Cols - 1
        h$ = h$ + vbTab
    Next i%
Else
    h$ = tRec!Name + vbTab + tRec!menge + vbTab + tRec!einheit
End If

MachTaxeAuswahlName$ = h$

Call clsError.DefErrPop
End Function

Function SuchWeiter%(ind%, Typ%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchWeiter%")
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
Dim aus%
Dim AltRecno&
Dim h$, AltKey$
Dim i%, AnzDS&, AnzDS2&
Dim lptr&, vgl$
Dim SQLStr$, SortierName$
Dim iVariant As Variant

SuchWeiter% = False

aus = 0
'Select Case iMatchModus%
'    Case TAXE_ALPHA, TAXE_MATCH
        MatchTaxeRec.Bookmark = Ausgabe(ind%).Bookmark
        
'        Do
'            If (Typ%) Then
'                MatchTaxeRec.MoveNext
'                If (MatchTaxeRec.EOF) Then
'                    Exit Do
'                End If
'            Else
'                MatchTaxeRec.MovePrevious
'                If (MatchTaxeRec.BOF) Then
'                    Exit Do
'                End If
'            End If
'            If (CheckOk(aus)) Then
'                SuchWeiter% = True
'                Exit Do
'            End If
'            If (aus) Then
'                Exit Do
'            End If
'        Loop
''''''''''''''''
        SortierName = Ausgabe(ind).key
        If (Typ%) Then
            MatchTaxeRec.MoveNext
            If (MatchTaxeRec.EOF) Then
                SortierName = Ausgabe(0).key
        '                    SQLStr$ = "SELECT COUNT(*) AS AnzDS FROM TAXE WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
                SQLStr = CountStr + " AND (SortierName>'" + SortierName + "')"
        '                    SQLStr = SQLStr + " ORDER BY SortNameNeu"
                AnzDS = 0
                MatchTaxeRec2.Open SQLStr$, MatchConn ' TaxeAdoDB1.ActiveConn
                If Not (MatchTaxeRec2.EOF) Then
                    AnzDS = clsOpTool.CheckNullLong(MatchTaxeRec2!AnzDS)
                End If
                MatchTaxeRec2.Close
                
                If (AnzDS >= MaxAnzeige - 1) Then
                    MatchTaxeRec.Close
        
                    SQLStr = SelectStr2 + " AND (SortierName<='" + SortierName + "')"
                    SQLStr = SQLStr + " ORDER BY SortierName DESC"
                    MatchTaxeRec.Open SQLStr$, MatchConn    'TaxeAdoDB1.ActiveConn
                        
                    MatchTaxeRec.MoveLast
                    SortierName = MatchTaxeRec!SortierName
                    MatchTaxeRec.Close
                    SQLStr = SelectStr + " AND (SortierName>='" + SortierName + "')"
                    SQLStr = SQLStr + " ORDER BY SortierName"
                    MatchTaxeRec.Open SQLStr$, MatchConn    'TaxeAdoDB1.ActiveConn
                    MatchTaxeRec.MoveLast
                        
                    Do
                        MatchTaxeRec.MovePrevious
                        If (Trim(MatchTaxeRec!SortierName) = Trim(Ausgabe(0).key)) Then
                            Exit Do
                        End If
                    Loop
        
                    Call Umspeichern(buf$, 0)
                    Call frmMatchcode.MachMaske
                    MatchTaxeRec.MoveNext
                    SuchWeiter% = True
                End If
            Else
                SuchWeiter% = True
            End If
        Else
            If (Trim(Ausgabe(0).key) = ErsterSortierName) Then
            Else
                MatchTaxeRec.MovePrevious
                If (MatchTaxeRec.BOF) Then
                    SortierName = Ausgabe(AnzAnzeige - 1).key
        '                        SQLStr$ = "SELECT COUNT(*) AS AnzDS FROM TAXE WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
                    SQLStr = CountStr + " AND (SortierName<'" + SortierName + "')"
        '                    SQLStr = SQLStr + " ORDER BY SortNameNeu"
                    AnzDS = 0
                    MatchTaxeRec2.Open SQLStr$, MatchConn   'TaxeAdoDB1.ActiveConn
                    If Not (MatchTaxeRec2.EOF) Then
                        AnzDS = clsOpTool.CheckNullLong(MatchTaxeRec2!AnzDS)
                    End If
                    MatchTaxeRec2.Close
                    
                    If (AnzDS >= MaxAnzeige - 1) Then
                        MatchTaxeRec.Close
        
        '                            SQLStr$ = "SELECT TOP(" + CStr((MaxRowsArbeit%(1) - 1) * 4) + ")* FROM TAXE WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
                        SQLStr = SelectStr2 + " AND (SortierName<='" + SortierName + "')"
                        SQLStr = SQLStr + " ORDER BY SortierName DESC"
                        MatchTaxeRec.Open SQLStr$, MatchConn    'TaxeAdoDB1.ActiveConn
                        
'                        AnzDS2 = (MaxAnzeige - 1) * 10
'                        If (AnzDS < AnzDS2) Then
'                            AnzDS2 = AnzDS
'                        End If
        
                        MatchTaxeRec.MoveLast
                        SortierName = MatchTaxeRec!SortierName
                        MatchTaxeRec.Close
                        SQLStr = SelectStr + " AND (SortierName>='" + SortierName + "')"
                        SQLStr = SQLStr + " ORDER BY SortierName"
                        MatchTaxeRec.Open SQLStr$, MatchConn    'TaxeAdoDB1.ActiveConn
                        MatchTaxeRec.MoveLast
'                        For i = 1 To (AnzAnzeige - 1)
'                            MatchTaxeRec.MovePrevious
'                        Next i
                        Do
                            MatchTaxeRec.MovePrevious
                            If (Trim(MatchTaxeRec!SortierName) = Trim(Ausgabe(0).key)) Then
                                Exit Do
                            End If
                        Loop
    '                    TaxeRec.Bookmark = AnzDS2 - AnzAnzeige
                        Call Umspeichern(buf$, 0)
                        Call frmMatchcode.MachMaske
                        MatchTaxeRec.Bookmark = Ausgabe(0).Bookmark
                        MatchTaxeRec.MovePrevious
                        SuchWeiter% = True
                    End If
                Else
                    SuchWeiter% = True
                End If
            End If
        End If
    

Call clsError.DefErrPop
End Function

Function SuchArtikel%(Such$, iArbeitAnzzeilen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchArtikel%")
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
Dim i%, l%, l2%, ind%, TextSuche%, IstPzn%, LeftLen%, InhMengeFlag%, gef%, IstCode%
Dim InhMenge&, AnzDS&
Dim SQLStr$, nSuch$, ch$, oMatchModus%
Dim KurzCode As Boolean
Dim iBookmark As Variant

Dim ActiveMatch%, LastErrf%, aus%
Dim txt$, OrgMatch$
Dim TaxeBookmark As Variant
Dim AltRecno&
Dim AltKey$

If (Trim(Such) = "") Then
    Call frmMatchcode.ZeigeInfoBereichAdd(0)
    frmMatchcode.txtMatchcode.SetFocus
    SuchArtikel = False
    Call clsError.DefErrPop: Exit Function
End If

SuchArtikel% = True

If (iMatchModus% = LAGER_MATCH) Or (iMatchModus% = MERKZETTEL_MATCH) Then
    Set MatchConn = ArtikelDB1.ActiveConn
    MatchDB = "Artikel"
    SqlServerDB = ArtikelDB1.SqlServerDB
Else
    Set MatchConn = TaxeAdoDB1.ActiveConn
    MatchDB = "Taxe"
    SqlServerDB = TaxeAdoDB1.SqlServerDB
End If

iOrgSuch$ = Such$

KurzCode = (Left$(Such$, 1) = "#")
If (KurzCode) Then
    Such$ = Mid$(Such$, 2)
End If
'
If (iMatchAutoRet%) And (Left$(Such$, 3) = "PZN") Then
    nSuch$ = Mid$(Such$, 4, 8)
    Such$ = Mid$(Such$, 12)
    iOrgSuch$ = Such$
    frmMatchcode.txtMatchcode.text = Such$

    On Error Resume Next
    MatchTaxeRec.Close
    On Error GoTo 0
    
    On Error Resume Next
    MatchTaxeRec2.Close
    On Error GoTo DefErr

    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + nSuch
    MatchTaxeRec.Open SQLStr$, TaxeAdoDB1.ActiveConn
    If (MatchTaxeRec.EOF = False) Then
        iMatchModus% = TAXE_ALPHA
        ErsterSortierName = MatchTaxeRec!SortierName
        
        MatchTaxeRec.Close
        
        If (TaxeAdoDB1.SqlServerDB) Then
            SelectStr$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 20) + ")* FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
            SelectStr2$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 10) + ")* FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        Else
            SelectStr$ = "SELECT TOP " + CStr((MaxAnzeige - 1) * 20) + " * FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        End If
        CountStr$ = "SELECT COUNT(*) AS AnzDS FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        SQLStr = SelectStr + " ORDER BY SortierName"
        
        ErsterSortierName = ""
        If (TaxeAdoDB1.SqlServerDB) Then
            SelectStr$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 20) + ")* FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
            SelectStr2$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 10) + ")* FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        Else
            SelectStr$ = "SELECT TOP " + CStr((MaxAnzeige - 1) * 20) + " * FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        End If
        CountStr$ = "SELECT COUNT(*) AS AnzDS FROM TAXE WHERE (SortierName>='" + ErsterSortierName + "')"
        
        MatchTaxeRec.Open SQLStr$, TaxeAdoDB1.ActiveConn
        Call SuchArtikelTrue(0)
        WarteMaske = 0
        Call clsError.DefErrPop: Exit Function
    End If

'    Set MatchTaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
'    MatchTaxeRec.Index = "Pzn"
'    MatchTaxeRec.Seek "=", nSuch$
'    If (MatchTaxeRec.NoMatch = False) Then
'        iBookmark = MatchTaxeRec.Bookmark
'        MatchTaxeRec.Index = "Name"
'        MatchTaxeRec.Bookmark = iBookmark
'    End If
'    If (MatchTaxeRec.NoMatch = False) Then
'        Call SuchArtikelTrue(False)
'        Call clsError.DefErrPop: Exit Function
'    End If
End If


If (Such$ = "") Then Such$ = Space$(10)
l% = Len(Such$)
If ((l = 8) Or (l% = 9)) And (Left$(Such$, 1) = "-") Then Such$ = Mid$(Such$, 2)
If ((l% >= 10) And (((Left$(Such$, 1) = "*") And (Mid$(Such$, l%, 1) = "*")) Or (InStr(Such$, "PZN") > 0))) Then
    nSuch$ = ""
    For i% = 1 To l%
        ch$ = Mid$(Such$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            nSuch$ = nSuch$ + ch$
        End If
    Next i%
    Such$ = nSuch$
End If

l% = Len(Trim(Such$))

IstCode% = False
If (l% >= 2) Or (KurzCode) Then
    IstCode% = True
    For i% = 1 To l%
        ch$ = Mid$(Such$, i%, 1)
        If (InStr("0123456789", ch$) <= 0) Then
            IstCode% = False
            Exit For
        End If
    Next i%
End If

'If (IstCode%) And (l% > 7) Then
'    Such$ = Right$(Space$(13) + Such$, 13)
'    FabsErrf% = Ass1.IndexSearch(1, Such$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        Ass1.GetRecord (FabsRecno& + 1)
'        Such$ = Ass1.pzn
'    Else
'        Such$ = Mid$(Such$, 6, 7)
'    End If
'End If

If (IstCode%) Then
    nSuch$ = Right$(Space$(13) + Such$, 13)
        
    On Error Resume Next
    MatchTaxeRec.Close
    On Error GoTo DefErr
    
    SQLStr$ = "SELECT * FROM Artikel WHERE Ean = '" + nSuch + "'"
    MatchTaxeRec.Open SQLStr$, ArtikelDB1.ActiveConn
    If (MatchTaxeRec.EOF) Then
        On Error Resume Next
        MatchTaxeRec.Close
        On Error GoTo DefErr
        
        If (Len(Such) = 13) Then
'            If (Left(Such, 4) = "9990") Then
'                Such = Mid(Such, 5, 8)
'            Else
'                Such = Mid$(Such, 5, 8)
'            End If
            Such = Mid(Such, 5, 8)
        ElseIf (Len(Such) = 8) Then
            If (clsOpTool.MakePzn(Such) <> Such) Then
                Such = "0" + Left(Such, 7)
            End If
        End If
        
        SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + clsOpTool.SqlPzn(Such$)
        MatchTaxeRec.Open SQLStr$, ArtikelDB1.ActiveConn
    Else
        Such = clsOpTool.PznString(MatchTaxeRec!pzn)
    End If
    If (MatchTaxeRec.EOF = False) Then
        If (iMatchNurEiner% = False) Then
            iMatchModus% = LAGER_MATCH
            Call SuchArtikelTrue(True)
        End If
    Else
        On Error Resume Next
        MatchTaxeRec.Close
        On Error GoTo DefErr
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(Such$)
        MatchTaxeRec.Open SQLStr$, TaxeAdoDB1.ActiveConn
        If (MatchTaxeRec.EOF = False) Then
            If (iMatchNurEiner% = False) Then
                iMatchModus% = TAXE_ALPHA
                Call SuchArtikelTrue(True)
            End If
        Else
            SuchArtikel% = False
        End If
    End If
    iMatchModus% = oMatchModus%
    
    If (SuchArtikel% And iMatchNurEiner%) Then
        iMatchcodePzn$ = clsOpTool.PznString(Val(Such$))
        iMatchcodeErg$ = iMatchcodePzn$
        Unload frmMatchcode
    End If
    
    Call clsError.DefErrPop: Exit Function
End If


ActiveMatch = 0
If (iMatchModus% = LAGER_MATCH) Then
    ActiveMatch = 2
End If

txt = Such
Packung = 0
SuchNGr = ""
herst = ""
'  AuswahlKZ = ""
MusterName = ""
Muster = ""
MusterSuche = False
MusterAnf = ""
OrgMatch = txt
l% = Len(OrgMatch)
If Right(Trim(txt), 1) = "+" Then
    ActiveMatch = 0
    txt = Trim(txt)
    txt = Left(txt, Len(txt) - 1)
    l% = Len(txt)
End If

For i = 1 To Len(txt)
    If InStr(",#+*", Mid(txt, i, 1)) > 0 Then
        MusterSucheBreak = False
        MusterSuche = True
        Muster = RTrim(Mid(txt, i))
        txt = Left(txt, i - 1)
        MusterAnf = UCase(txt)
        Exit For
    End If
Next i

If ActiveMatch = 0 Then
    txt = UCase(txt)
    Call MakeTaxeSuch(txt, InhMenge&)
ElseIf ActiveMatch = 1 Then
    txt = MakeTaxeMatchcode(txt)
Else
    txt = MakeMatchCode$(txt)
End If
SqlTaxeSuch = txt
    
If MusterSuche Then
    l% = Len(txt)
    Call ZerlegeMuster(Muster, Packung, SuchNGr, herst)
    WarteMaske = True
End If
  
'..............................................................................
On Error Resume Next
MatchTaxeRec.Close
On Error GoTo 0

On Error Resume Next
MatchTaxeRec2.Close
On Error GoTo DefErr

    If (MusterSuche) Then
        SQLStr = "Komprimiert LIKE '" + txt$ + "%"
        If (Muster <> "") Then
            SQLStr = SQLStr + Muster + "%"
        End If
        SQLStr = SQLStr + "'"
        If (Packung > 0) Then
'            SQLStr = SQLStr + " AND (val(menge)=" + CStr(Packung) + ")"
            SQLStr = SQLStr + " AND (StdMenge=" + CStr(Packung) + ")"
        End If
        If (SuchNGr > "") Then
            SQLStr = SQLStr + " AND (NGrösse='" + SuchNGr + "')"
        End If
        If (herst > "") Then
            SQLStr = SQLStr + " AND (Left(HerstellerKB," + CStr(Len(herst)) + ") LIKE '" + herst + "')"
        End If
        
        If (SqlServerDB) Then
            SelectStr$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 20) + ")* FROM " + MatchDB + " WHERE " + SQLStr
            SelectStr2$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 10) + ")* FROM " + MatchDB + " WHERE " + SQLStr
        Else
            SelectStr$ = "SELECT TOP " + CStr((MaxAnzeige - 1) * 20) + " * FROM " + MatchDB + " WHERE " + SQLStr
        End If
        If (iMatchModus% = LAGER_MATCH) Then
            SelectStr = SelectStr + " AND LagerKz<>0"
            SelectStr2 = SelectStr2 + " AND LagerKz<>0"
        End If
        CountStr$ = "SELECT COUNT(*) AS AnzDS FROM " + MatchDB + " WHERE " + SQLStr
    Else
        If (SqlServerDB) Then
            SelectStr$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 20) + ")* FROM " + MatchDB + " WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
            SelectStr2$ = "SELECT TOP(" + CStr((MaxAnzeige - 1) * 10) + ")* FROM " + MatchDB + " WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
        Else
            SelectStr$ = "SELECT TOP " + CStr((MaxAnzeige - 1) * 20) + " * FROM " + MatchDB + " WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
        End If
        CountStr$ = "SELECT COUNT(*) AS AnzDS FROM " + MatchDB + " WHERE Komprimiert LIKE '" + SqlTaxeSuch + "%'"
    End If
    If (iMatchModus% = LAGER_MATCH) Then
        SelectStr = SelectStr + " AND LagerKz<>0"
        SelectStr2 = SelectStr2 + " AND LagerKz<>0"
        CountStr = CountStr + " AND LagerKz<>0"
    End If
    
    AnzDS = 0
    MatchTaxeRec2.Open CountStr$, MatchConn ' TaxeAdoDB1.ActiveConn
    If Not (MatchTaxeRec2.EOF) Then
        AnzDS = clsOpTool.CheckNullLong(MatchTaxeRec2!AnzDS)
    End If
    MatchTaxeRec2.Close
    
    If (AnzDS > 0) Then
        SQLStr = SelectStr + " ORDER BY SortierName"
'        SQLStr = SQLStr + " ORDER BY SortNr"
        MatchTaxeRec.Open SQLStr$, MatchConn    'TaxeAdoDB1.ActiveConn
        If (MatchTaxeRec.EOF) Then
            SuchArtikel% = False
        Else
            ErsterSortierName = MatchTaxeRec!SortierName
'            Call Umspeichern("", 0)
'            ListeSize = AnzDS  'MaxRowsArbeit%(1) - 1
'            AuswahlTyp% = liste 'ALLGEMEIN
        End If
    Else
        SuchArtikel% = False
    End If
'End If

IstPzn = 0
If (SuchArtikel% = True) Then
    Call SuchArtikelTrue(IstPzn%)
End If
WarteMaske = 0

Call clsError.DefErrPop
End Function

'Sub SuchArtikelTrue(Optional PznFlag% = False)
Sub SuchArtikelTrue(PznFlag%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SuchArtikelTrue")
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
Dim j%, ind%
Dim h$

Call Umspeichern("", 0)
Call frmMatchcode.MachAuswahlGrid
If (PznFlag%) Then
    h$ = Ausgabe(0).Name
    With frmMatchcode.flxarbeit(0)
        .Rows = 2
        For j% = 0 To 2
            ind% = InStr(h$, vbTab)
            .TextMatrix(1, j%) = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Next j%
        .TextMatrix(1, 3) = h$
        .row = 1
    End With
    Call ListeAusFlexBefuellen
    With frmMatchcode.txtMatchcode
        .SelStart = 0
        .SelLength = Len(.text)
    End With
Else
    With frmMatchcode
        Call .MachMaske
        .flxarbeit(0).Visible = True
        .txtFlexBack.Visible = True
        .flxarbeit(1).Visible = False
    End With
    With frmMatchcode.flxarbeit(0)
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
        .SetFocus
    End With
End If

Call clsError.DefErrPop
End Sub

Sub MachAuswahlGrid(iArbeitAnzzeilen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachAuswahlGrid")
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
Dim spBreite%, i%
Dim aFont$

With frmMatchcode.flxarbeit(0)
    .Rows = iArbeitAnzzeilen%
    .FixedRows = 1
            
    Select Case TaxeAuswahlModus%
        Case LANGERNAME
            .Cols = 6
            .FormatString = "|<Name|>Menge|<Eh|Zu|<Kz"
            frmMatchcode.Font.Bold = True
            .ColWidth(0) = frmMatchcode.TextWidth("X")
            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
            .ColWidth(4) = frmMatchcode.TextWidth("XX")
            .ColWidth(5) = frmMatchcode.TextWidth("XXX")
            frmMatchcode.Font.Bold = False
        Case AEPundAVP, PZNundAVP
            If (iMatchModus% = LAGER_MATCH) Or (iMatchModus% = MERKZETTEL_MATCH) Then
                .Cols = 9   ' 16
                .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
                frmMatchcode.Font.Bold = True
                .ColWidth(0) = frmMatchcode.TextWidth("X")
                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
                .ColWidth(4) = frmMatchcode.TextWidth("XX")
                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
                aFont$ = frmMatchcode.Font.Name
                frmMatchcode.Font.Name = "Courier New"
                frmMatchcode.Font.Bold = True
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
                frmMatchcode.Font.Name = aFont$
                frmMatchcode.Font.Bold = False
            Else
                .Cols = 9
                If (TaxeAuswahlModus% = AEPundAVP) Then
                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
                Else
                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|<PZN|>VK|<Kz"
                End If
                frmMatchcode.Font.Bold = True
                .ColWidth(0) = frmMatchcode.TextWidth("X")
                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
                .ColWidth(4) = frmMatchcode.TextWidth("XX")
                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXX")
                frmMatchcode.Font.Bold = False
            End If
        Case FESTBETRAG
            .Cols = 9
            .FormatString = "|<Name|>Menge|<Eh|^Zu|>FB|>DIFF|>VK|<Kz"
            frmMatchcode.Font.Bold = True
            .ColWidth(0) = frmMatchcode.TextWidth("X")
            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
            .ColWidth(4) = frmMatchcode.TextWidth("XX")
            .ColWidth(5) = frmMatchcode.TextWidth("99999.99")
            .ColWidth(6) = frmMatchcode.TextWidth("999.99")
            .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
            .ColWidth(8) = frmMatchcode.TextWidth("XXX")
            frmMatchcode.Font.Bold = False
    End Select
            
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (i <> 1) Then
            .ColWidth(i%) = .ColWidth(i%) + frmMatchcode.TextWidth("X")
            spBreite% = spBreite% + .ColWidth(i%)
        End If
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(1) = .Width - spBreite%
End With
Call clsError.DefErrPop
End Sub

Sub InitAuswahlGrid()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitAuswahlGrid")
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
Dim spBreite%, i%
Dim aFont$

With frmMatchcode.flxarbeit(0)
            
    .Cols = 4
    .FormatString = "|<Name|>Menge|<Eh"
    frmMatchcode.Font.Bold = True
    .ColWidth(0) = frmMatchcode.TextWidth("XX")
    .ColWidth(2) = frmMatchcode.TextWidth("XXXXXXX")
    .ColWidth(3) = frmMatchcode.TextWidth("XXXX")
    frmMatchcode.Font.Bold = False
    
'    Select Case TaxeAuswahlModus%
'        Case LANGERNAME
'            .Cols = 6
'            .FormatString = "|<Name|>Menge|<Eh|Zu|<Kz"
'            frmMatchcode.Font.Bold = True
'            .ColWidth(0) = frmMatchcode.TextWidth("X")
'            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
'            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
'            .ColWidth(4) = frmMatchcode.TextWidth("XX")
'            .ColWidth(5) = frmMatchcode.TextWidth("XXX")
'            frmMatchcode.Font.Bold = False
'        Case AEPundAVP, PZNundAVP
'            If (MatchModus% = LAGER_MATCH) Then
'                .Cols = 9   ' 16
'                .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
'                frmMatchcode.Font.Bold = True
'                .ColWidth(0) = frmMatchcode.TextWidth("X")
'                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
'                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
'                .ColWidth(4) = frmMatchcode.TextWidth("XX")
'                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
'                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
'                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
'                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
'                aFont$ = frmMatchcode.Font.Name
'                frmMatchcode.Font.Name = "Courier New"
'                frmMatchcode.Font.Bold = True
'                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXXXXXX99XX")
'                frmMatchcode.Font.Name = aFont$
'                frmMatchcode.Font.Bold = False
'            Else
'                .Cols = 9
'                If (TaxeAuswahlModus% = AEPundAVP) Then
'                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|>EK|>VK|<Kz"
'                Else
'                    .FormatString = "|<Name|>Menge|<Eh|^Zu|<Herst|<PZN|>VK|<Kz"
'                End If
'                frmMatchcode.Font.Bold = True
'                .ColWidth(0) = frmMatchcode.TextWidth("X")
'                .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
'                .ColWidth(3) = frmMatchcode.TextWidth("XXX")
'                .ColWidth(4) = frmMatchcode.TextWidth("XX")
'                .ColWidth(5) = frmMatchcode.TextWidth("XXXXXX")
'                .ColWidth(6) = frmMatchcode.TextWidth("99999.99")
'                .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
'                .ColWidth(8) = frmMatchcode.TextWidth("XXXXXX")
'                frmMatchcode.Font.Bold = False
'            End If
'        Case FESTBETRAG
'            .Cols = 9
'            .FormatString = "|<Name|>Menge|<Eh|^Zu|>FB|>DIFF|>VK|<Kz"
'            frmMatchcode.Font.Bold = True
'            .ColWidth(0) = frmMatchcode.TextWidth("X")
'            .ColWidth(2) = frmMatchcode.TextWidth("XXXXXX")
'            .ColWidth(3) = frmMatchcode.TextWidth("XXX")
'            .ColWidth(4) = frmMatchcode.TextWidth("XX")
'            .ColWidth(5) = frmMatchcode.TextWidth("99999.99")
'            .ColWidth(6) = frmMatchcode.TextWidth("999.99")
'            .ColWidth(7) = frmMatchcode.TextWidth("99999.99")
'            .ColWidth(8) = frmMatchcode.TextWidth("XXX")
'            frmMatchcode.Font.Bold = False
'    End Select
            
'    spBreite% = 0
'    For i% = 0 To .Cols - 1
'        If (i <> 1) Then
'            .ColWidth(i%) = .ColWidth(i%) + frmMatchcode.TextWidth("X")
'            spBreite% = spBreite% + .ColWidth(i%)
'        End If
'    Next i%
'    If (spBreite% > .Width) Then
'        spBreite% = .Width
'    End If
'    .ColWidth(1) = .Width - spBreite%
End With
Call clsError.DefErrPop
End Sub

Sub ListeAusFlexBefuellen(Optional AlphaText$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeAusFlexBefuellen")
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
Dim row%
Dim pzn$
        
If (AlphaText$ <> "") Then
    Call ListeBefuellen(clsOpTool.PznString(9999999), AlphaText$, "", "")
Else
    
    With frmMatchcode.flxarbeit(0)
        row% = .row
        
        pzn$ = clsOpTool.PznString(Ausgabe(row% - 1).pzn)
        Call ListeBefuellen(pzn$, .TextMatrix(row%, 1), .TextMatrix(row%, 2), .TextMatrix(row%, 3))
    End With
End If

Call clsError.DefErrPop
End Sub

Sub ListeAusStiftBefuellen(Dp04Str$, Dp04Menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeAusStiftBefuellen")
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
Dim row%, iLen%
Dim h$, pzn$, SQLStr$, AlphaText$, Such$, nSuch$
        
AlphaText$ = ""

If (ArtikelDBok) Then
    Such = Dp04Str
    nSuch$ = Right$(Space$(13) + Such, 13)
        
    On Error Resume Next
    MatchTaxeRec.Close
    On Error GoTo DefErr
    
    SQLStr$ = "SELECT * FROM Artikel WHERE Ean = '" + nSuch + "'"
    MatchTaxeRec.Open SQLStr$, ArtikelDB1.ActiveConn
    If (MatchTaxeRec.EOF) Then
        On Error Resume Next
        MatchTaxeRec.Close
        On Error GoTo DefErr
        
        If (Len(Such) = 13) Then
            If (Left(Such, 4) = "9990") Then
                Such = Mid(Such, 5, 8)
            End If
        ElseIf (Len(Such) = 8) Then
            If (clsOpTool.MakePzn(Such) <> Such) Then
                Such = "0" + Left(Such, 7)
            End If
        End If
        
        SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + clsOpTool.SqlPzn(Such$)
        MatchTaxeRec.Open SQLStr$, ArtikelDB1.ActiveConn
    End If
    If (MatchTaxeRec.EOF = False) Then
        pzn = clsOpTool.PznString(MatchTaxeRec!pzn)
    Else
        On Error Resume Next
        MatchTaxeRec.Close
        On Error GoTo DefErr
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(Such$)
        MatchTaxeRec.Open SQLStr$, TaxeAdoDB1.ActiveConn
        If (MatchTaxeRec.EOF = False) Then
            pzn = Such
        Else
            pzn$ = clsOpTool.PznString(9999999)
            AlphaText$ = "Artikel unbekannt"
        End If
    End If
Else
    h$ = Dp04Str$
    h$ = Right$(Space$(13) + h$, 13)
    
    FabsErrf% = Ass1.IndexSearch(1, h$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Ass1.GetRecord (FabsRecno& + 1)
        pzn$ = Ass1.pzn
    Else
        If (Len(Dp04Str$) = 7) Then
            h$ = Dp04Str$
        Else
            h$ = Mid$(h$, 6, 7)
        End If
        FabsErrf% = Ass1.IndexSearch(0, h$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Ass1.GetRecord (FabsRecno& + 1)
            pzn$ = Ass1.pzn
        Else
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(h$)
            On Error Resume Next
            MatchTaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            MatchTaxeRec.Open SQLStr, TaxeAdoDB1.ActiveConn
            If (MatchTaxeRec.EOF = False) Then
                pzn$ = h$
            Else
                pzn$ = "9999999"
                AlphaText$ = "Artikel unbekannt"
            End If
        End If
    End If
End If

Call ListeBefuellen(pzn$, AlphaText$, "", "", Dp04Menge%)

Call clsError.DefErrPop
End Sub
  
Sub ListeBefuellen(pzn$, txt1$, txt2$, txt3$, Optional Dp04Menge% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ListeBefuellen")
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
Dim i%, erg%, row%, ManuellAsatz%, ManuellSsatz%, ManuellBm%, ManuellNm%, abl%, ind%, iMenge%, OrgWi%, iLiefNr%, iLiefDat%
Dim h$, h2$, AufLager$, SQLStr$
        
iMenge% = 1
If (Dp04Menge% > 0) Then iMenge% = Dp04Menge%

If (Val(pzn$) = 9999999) Then
    h$ = "09999999" + vbTab + txt1$ + vbTab + vbTab + vbTab
    h$ = h$ + Str$(iMenge%) + vbTab + "0"
Else
    ManuellAsatz% = 0
    ManuellSsatz% = 0
    
    If (ArtikelDBok) Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
    '    SQLStr = SQLStr + " AND Lagerkz<>0"
        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
        If (FabsErrf = 0) Then
            ManuellAsatz% = 1
            ManuellSsatz% = 1
        
            If (iMatchListeTyp% = 2) And (Ass1.poslag > 0) Then
                iMenge% = Ass1.poslag
            End If
        Else
            Call clsOpTool.Taxe2ast(pzn$)
        End If
    
        If (Trim(txt1$) = "") Then
            txt1$ = Ast1.kurz
            txt2$ = Ast1.meng
            txt3$ = Ast1.meh
        End If
    Else
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ManuellAsatz% = CInt(FabsRecno&)
            Ast1.GetRecord (FabsRecno& + 1)
        End If
        If (ManuellAsatz% = 0) Then
            Call clsOpTool.Taxe2ast(pzn$)
        End If
                
        If (Trim(txt1$) = "") Then
            txt1$ = Ast1.kurz
            txt2$ = Ast1.meng
            txt3$ = Ast1.meh
        End If
        
        FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ManuellSsatz% = CInt(FabsRecno&)
            Ass1.GetRecord (FabsRecno& + 1)
            If (iMatchListeTyp% = 2) And (Ass1.poslag > 0) Then
                iMenge% = Ass1.poslag
            End If
        End If
    End If
    
    If (iMatchListeTyp%) Then
        h$ = pzn$ + vbTab + Trim$(txt1$) + vbTab + Trim$(txt2$) + vbTab + Trim$(txt3$)
        
        h$ = h$ + vbTab + ""
        h$ = h$ + vbTab + ""
        h$ = h$ + vbTab + ""
        
        h$ = h$ + vbTab + Str$(iMenge%)
        
        h2$ = ""
        If (ManuellSsatz% > 0) Then
            If (Ass1.poslag > 0) Then
                abl% = Ass1.abl2
                If (Ass1.abl1 > abl%) Then abl% = Ass1.abl1
                If (abl% > 0) Then
                    h2$ = Right$(clsOpTool.sDate(abl%), 4)
                End If
            End If
        End If
        h$ = h$ + vbTab + h2$

        h2$ = Format(Ast1.aep, "0.00")
        h$ = h$ + vbTab + h2$
    Else
        If (Dp04Menge% > 0) Then
            ManuellBm% = Dp04Menge%
        ElseIf (ManuellSsatz% > 0) Then
            ManuellBm% = Ass1.bm
            If (Ass1.vbm > 0) Then ManuellBm% = Ass1.vbm
        Else
            ManuellBm% = 1
            ManuellNm% = 0
        End If
        ' neu am 9.1.13
        If (ManuellBm = 0) Then
            ManuellBm = 1
        End If
    
        AufLager$ = ""
        If (ManuellSsatz% = 0) Then
            If (Val(pzn$) = 9999999) Or (ManuellAsatz% = 0) Then
                AufLager$ = "-"
'                Else
'                    erg% = MsgBox("Ist dieser Artikel bereits auf Lager ?", vbYesNo Or vbInformation)
'                    If (erg% = vbNo) Then
'                        AufLager$ = "-"
'                    End If
            End If
        End If
    
        h$ = pzn$ + vbTab + Trim$(txt1$) + vbTab + Trim$(txt2$) + vbTab + Trim$(txt3$)
        h$ = h$ + vbTab + Str$(ManuellBm%) + vbTab + Str$(ManuellNm%)
        h$ = h$ + vbTab + AufLager$
    End If
End If

With frmMatchcode.flxarbeit(1)
    .AddItem h$, 1
    .row = 1
'    .AddItem h$
'    .row = .Rows - 1
    .col = 3
    .col = 4
    .ColSel = 4
    
    .Visible = True
    frmMatchcode.flxarbeit(0).Visible = False
    frmMatchcode.txtFlexBack.Visible = False
    
    If (iMatchListeTyp% = 1) Then
        If (clsOpTool.SecurPharmDMC <> "") Then
            iLiefNr = clsOpTool.SecurPharmLiefNr
            h2$ = ""
            If (LieferantenDBok) Then
                h = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(iLiefNr)
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                End If
                LieferantenRec.Close
            Else
                Call Lif1.GetRecord(iLiefNr + 1)
                h$ = Lif1.kurz
                h$ = Trim$(h$)
            End If
            If (h$ <> "") Then
                If (Asc(Left$(h$, 1)) >= 32) Then
                    h$ = h$ + " (" + Mid$(Str$(iLiefNr), 2) + ")"
                    h2$ = h$
                End If
            End If
            .TextMatrix(.row, 4) = h2
            
            .TextMatrix(.row, 5) = clsOpTool.SecurPharmBeleg
            
            h = ""
            iLiefDat = clsOpTool.SecurPharmLieferDatum
            If (iLiefDat > 0) Then
                h = clsOpTool.sDate(iLiefDat)
            End If
            .TextMatrix(.row, 6) = h
            
            .TextMatrix(.row, 7) = "1"
'            .TextMatrix(.row, 8) = Mid(clsOpTool.SecurPharmValue(3), 3)
            h = Trim(clsOpTool.SecurPharmValue(3))
            If (h <> "") Then
                h = Mid(h, 3, 2) + Left(h, 2)
            End If
            .TextMatrix(.row, 8) = h

            
            .TextMatrix(.row, 10) = clsOpTool.SecurPharmValue(2)
        ElseIf (ArtikelDBok) Then
            SQLStr$ = "SELECT * FROM LIEFERUNGEN WHERE PZN = " + pzn
            SQLStr = SQLStr + " ORDER BY BelegDatum DESC"
            FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
            If (FabsErrf% = 0) Then
                With frmMatchcode
                    OrgWi = .flxarbeit(1).ColWidth(8)
                    .flxarbeit(1).ColWidth(1) = .flxarbeit(1).ColWidth(1) - 1.5 * OrgWi
                    .flxarbeit(1).ColWidth(8) = 2.5 * OrgWi
                    
                    frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(4) '+ 45
                    frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                    frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + (.flxarbeit(1).row - .flxarbeit(1).TopRow + 1) * .flxarbeit(1).RowHeight(0)
                    frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                    frmEdit2.Width = .flxarbeit(1).ColPos(10) - .flxarbeit(1).ColPos(4)
                    frmEdit2.Height = .flxarbeit(1).RowHeight(0) * 5 '+ 90
                End With
                
                With frmEdit2.flxEdit
                    .BorderStyle = flexBorderNone
                    .ScrollBars = flexScrollBarNone
                    
                    .Height = frmEdit2.ScaleHeight
                    frmEdit2.Height = .Height
                    .Width = frmEdit2.ScaleWidth
                    .Left = 0
                    .Top = 0
                    
                    .Rows = 0
                    .Cols = 6
                    
                    For i = 0 To 5
                        .ColWidth(i) = frmMatchcode.flxarbeit(1).ColWidth(i + 4)
                        .ColAlignment(i%) = flexAlignLeftCenter
                    Next i%
                    .ColAlignment(5) = flexAlignRightCenter
                    
                    For i% = 0 To 4
                        If (ArtikelRec2.EOF) Then
                            Exit For
                        End If
                        
                        iLiefNr = clsOpTool.CheckNullLong(ArtikelRec2!LiefNr)
                        h2$ = ""
                        If (LieferantenDBok) Then
                            h = ""
                            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(iLiefNr)
                            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                            If (LieferantenRec.RecordCount <> 0) Then
                                h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                            End If
                            LieferantenRec.Close
                        Else
                            Call Lif1.GetRecord(iLiefNr + 1)
                            h$ = Lif1.kurz
                            h$ = Trim$(h$)
                        End If
                        If (h$ <> "") Then
                            If (Asc(Left$(h$, 1)) >= 32) Then
                                h$ = h$ + " (" + Mid$(Str$(iLiefNr), 2) + ")"
                                h2$ = h$
                            End If
                        End If
                        
                        h$ = h2$ + vbTab + clsOpTool.CheckNullStr(ArtikelRec2!Beleg) + vbTab + Format(ArtikelRec2!BelegDatum, "DDMMYY") + vbTab
                        If (ArtikelRec2!RetMenge > 0) Then
                            iMenge = ArtikelRec2!lm - ArtikelRec2!RetMenge
                            If (iMenge < 0) Then
                                iMenge = 0
                            End If
                            h$ = h$ + Str$(iMenge) + vbTab + " (urspr.LM: " + Str$(ArtikelRec2!lm) + ")" + vbTab
                        Else
                            h$ = h$ + Str$(ArtikelRec2!lm) + vbTab + vbTab
                        End If
                        h$ = h$ + Format(ArtikelRec2!NNAEP, "0.00")
                        .AddItem h$
                        
                        ArtikelRec2.MoveNext
                    Next i%
                    
                    .row = 0
                    .col = 0
                    .ColSel = .Cols - 1
                    .Visible = True
                End With
                
                frmEdit2.Show 1
                    
                With frmMatchcode
                    .flxarbeit(1).ColWidth(1) = .flxarbeit(1).ColWidth(1) + 1.5 * OrgWi
                    .flxarbeit(1).ColWidth(8) = OrgWi
                End With
                
                If (EditErg%) Then
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 4) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
            
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 5) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
                    
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 6) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
                    
                    h$ = EditTxt$
                    .TextMatrix(.row, 7) = Trim$(h$)
                End If
            End If
        Else
    '        nnek1.OpenDatei
            FabsErrf% = nnek1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                nnek1.GetRecord (FabsRecno& + 1)
                
                With frmMatchcode
                    OrgWi = .flxarbeit(1).ColWidth(8)
                    .flxarbeit(1).ColWidth(1) = .flxarbeit(1).ColWidth(1) - 1.5 * OrgWi
                    .flxarbeit(1).ColWidth(8) = 2.5 * OrgWi
                    
                    frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(4) '+ 45
                    frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                    frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + (.flxarbeit(1).row - .flxarbeit(1).TopRow + 1) * .flxarbeit(1).RowHeight(0)
                    frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                    frmEdit2.Width = .flxarbeit(1).ColPos(10) - .flxarbeit(1).ColPos(4)
                    frmEdit2.Height = .flxarbeit(1).RowHeight(0) * 5 '+ 90
                End With
                
                With frmEdit2.flxEdit
                    .BorderStyle = flexBorderNone
                    .ScrollBars = flexScrollBarNone
                    
                    .Height = frmEdit2.ScaleHeight
                    frmEdit2.Height = .Height
                    .Width = frmEdit2.ScaleWidth
                    .Left = 0
                    .Top = 0
                    
                    .Rows = 0
                    .Cols = 6
                    
                    For i = 0 To 5
                        .ColWidth(i) = frmMatchcode.flxarbeit(1).ColWidth(i + 4)
                        .ColAlignment(i%) = flexAlignLeftCenter
                    Next i%
                    .ColAlignment(5) = flexAlignRightCenter
                    
                    For i% = 0 To 4
                        h2$ = ""
                        If (LieferantenDBok) Then
                            h = ""
                            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(nnek1.LiefNr(i%))
                            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                            If (LieferantenRec.RecordCount <> 0) Then
                                h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                            End If
                            LieferantenRec.Close
                        Else
                            Call Lif1.GetRecord(nnek1.LiefNr(i%) + 1)
                            h$ = Lif1.kurz
                            h$ = Trim$(h$)
                        End If
                        If (h$ <> "") Then
                            If (Asc(Left$(h$, 1)) >= 32) Then
                                h$ = h$ + " (" + Mid$(Str$(nnek1.LiefNr(i%)), 2) + ")"
                                h2$ = h$
                            End If
                        End If
                        
                        h$ = h2$ + vbTab + nnek1.Beleg(i%) + vbTab + clsOpTool.sDate(nnek1.BelegDat(i%)) + vbTab
                        If (nnek1.RetMenge(i) > 0) Then
                            iMenge = nnek1.lm(i%) - nnek1.RetMenge(i)
                            If (iMenge < 0) Then
                                iMenge = 0
                            End If
                            h$ = h$ + Str$(iMenge) + vbTab + " (urspr.LM: " + Str$(nnek1.lm(i%)) + ")" + vbTab
                        Else
                            h$ = h$ + Str$(nnek1.lm(i%)) + vbTab + vbTab
                        End If
                        h$ = h$ + Format(nnek1.NNAEP(i%), "0.00")
                        .AddItem h$
                    Next i%
                    
                    .row = 0
                    .col = 0
                    .ColSel = .Cols - 1
                    .Visible = True
                End With
                
                frmEdit2.Show 1
                    
                With frmMatchcode
                    .flxarbeit(1).ColWidth(1) = .flxarbeit(1).ColWidth(1) + 1.5 * OrgWi
                    .flxarbeit(1).ColWidth(8) = OrgWi
                End With
                
                If (EditErg%) Then
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 4) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
            
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 5) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
                    
                    ind% = InStr(EditTxt$, vbTab)
                    h$ = Left$(EditTxt$, ind% - 1)
                    .TextMatrix(.row, 6) = Trim$(h$)
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
                    
                    h$ = EditTxt$
                    .TextMatrix(.row, 7) = Trim$(h$)
                End If
            End If
    '        nnek1.CloseDatei
        End If
    End If
    
    frmMatchcode.txtMatchcode.SetFocus
 
End With

Call clsError.DefErrPop
End Sub

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, arow%, m%, aMeng%, iLief%
Dim h$, h2$
            
With frmMatchcode.flxarbeit(1)
    EditCol% = .col
    If ((EditCol% = 1) And (Val(.TextMatrix(.row, 0)) <> 9999999)) Or _
       ((iMatchListeTyp% = 0) And (EditCol% <> 4) And (EditCol% <> 5)) Or _
       ((iMatchListeTyp% > 0) And ((EditCol% = 2) Or (EditCol% = 3))) Then
        Call clsError.DefErrPop: Exit Sub
    End If
            
        
'    If ((EditCol% = 1) And (.TextMatrix(.row, 0) = "9999999")) Or (EditCol% = 4) Or (EditCol% = 5) Then
        
        KeinRowColChange% = True
        arow% = .row
        .row = 0
        .CellFontBold = True
        .row = arow%
        KeinRowColChange% = False
            
        Load frmEdit2
        
        If (iMatchListeTyp% = 1) And (EditCol% = 4) Then
            lInd% = -1
            With frmMatchcode
                frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(EditCol%) + 45
                frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + .flxarbeit(1).RowHeight(0)
                frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                frmEdit2.Width = .flxarbeit(1).ColWidth(EditCol%) * 1.5
                frmEdit2.Height = .flxarbeit(1).Height - .flxarbeit(1).RowHeight(0)
            End With
            With frmEdit2.lstEdit
                .Height = frmEdit2.ScaleHeight
                frmEdit2.Height = .Height
                .Width = frmEdit2.ScaleWidth
                .Left = 0
                .Top = 0
                
                .Clear
'                If (LieferantenDBok) Then
'                    For i% = 1 To LieferantenDB1.AnzLiefNamen
'                        h$ = LieferantenDB1.LiefName(i% - 1)
'                        .AddItem h$
'                    Next i%
'                Else
                    For i% = 1 To Lif1.AnzLiefNamen
                        h$ = Lif1.LiefName(i% - 1)
                        .AddItem h$
                    Next i%
'                End If

'                If (lInd% < 0) Then
'                    .ListIndex = 0
'                Else
'                    .ListIndex = lInd%
'                End If
                .Visible = True
                EditModus% = 1
            End With
            
        Else
        
            With frmMatchcode
                frmEdit2.Left = .picBack(0).Left + .flxarbeit(1).Left + .flxarbeit(1).ColPos(EditCol%) + 45
                frmEdit2.Left = frmEdit2.Left + .Left + wPara1.FrmBorderHeight
                frmEdit2.Top = .picBack(0).Top + .flxarbeit(1).Top + (.flxarbeit(1).row - .flxarbeit(1).TopRow + 1) * .flxarbeit(1).RowHeight(0)
                frmEdit2.Top = frmEdit2.Top + .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + wPara1.FrmMenuHeight
                frmEdit2.Width = .flxarbeit(1).ColWidth(EditCol%)
                frmEdit2.Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
            End With
            With frmEdit2.txtEdit
                .Width = frmEdit2.ScaleWidth
    '            .Height = frmEdit.ScaleHeight
                .Left = 0
                .Top = 0
                h2$ = frmMatchcode.flxarbeit(1).TextMatrix(frmMatchcode.flxarbeit(1).row, EditCol%)
                .text = h2$
                .BackColor = vbWhite
                .Visible = True
            End With
            
            EditModus% = 0
            If (EditCol% = 1) Then
                EditModus% = 1
            ElseIf (iMatchListeTyp% > 0) Then
                If (EditCol% = 5) Then
                    EditModus% = 1
                ElseIf (EditCol% = 6) Then
                    EditModus% = 6
                ElseIf (EditCol% = 8) Then
                    EditModus% = 5
                ElseIf (EditCol% = 9) Then
                    EditModus% = 4
                End If
            End If
        End If
            
        frmEdit2.Show 1
            
        KeinRowColChange% = True
        arow% = .row
        .row = 0
        .CellFontBold = False
        .row = arow%
        KeinRowColChange% = False
            
        If (EditErg%) Then
            If (EditModus% = 0) Then
                m% = Val(EditTxt$)
                .TextMatrix(.row, EditCol%) = Mid$(Str$(m%), 2)
            ElseIf (EditModus% = 1) Or (EditModus% = 5) Or (EditModus% = 6) Then
                .TextMatrix(.row, EditCol%) = Trim(EditTxt$)
            ElseIf (EditModus% = 4) Then
                .TextMatrix(.row, EditCol%) = Format(Val(EditTxt$), "0.00")
            End If
        End If
'    End If
End With

Call clsError.DefErrPop
End Sub
      

Public Property Get IniSection$()
If (iMatchListeTyp% = 1) Then
    IniSection$ = "Auswahl-Retouren"
ElseIf (iMatchListeTyp% = 2) Then
    IniSection$ = "Auswahl-Rückkauf-Artikel"
Else
    IniSection$ = INI_SECTION
End If
End Property

Public Property Get InfoSection$()
InfoSection$ = INFO_SECTION
End Property

Public Property Get ArbeitSection$()
ArbeitSection$ = ARBEIT_SECTION
End Property

Public Property Get OrgCols%()
OrgCols% = 4
End Property

Public Property Get EingabeName$()
EingabeName$ = EINGABE_NAME
End Property


Public Property Get MatchModus%()
MatchModus% = iMatchModus%
End Property
Public Property Let MatchModus(ByVal vNewValue%)
iMatchModus% = vNewValue%
End Property


Public Property Get MatchTyp%()
MatchTyp% = iMatchTyp%
End Property
Public Property Let MatchTyp(ByVal vNewValue%)
iMatchTyp% = vNewValue%
End Property


Public Property Get MatchListeTyp%()
MatchListeTyp% = iMatchListeTyp%
End Property
Public Property Let MatchListeTyp(ByVal vNewValue%)
iMatchListeTyp% = vNewValue%
End Property


Public Property Get MatchAnzeigeTyp$(ByVal ind%)
MatchAnzeigeTyp$ = iMatchAnzeigeTyp$(ind%)
End Property
Public Property Let MatchAnzeigeTyp(ByVal ind%, ByVal vNewValue$)
iMatchAnzeigeTyp$(ind%) = vNewValue$
End Property


Public Property Get MatchAutoRet%()
MatchAutoRet% = iMatchAutoRet%
End Property
Public Property Let MatchAutoRet(ByVal vNewValue%)
iMatchAutoRet% = vNewValue%
End Property


Public Property Get MatchNurEiner%()
MatchNurEiner% = iMatchNurEiner%
End Property
Public Property Let MatchNurEiner(ByVal vNewValue%)
iMatchNurEiner% = vNewValue%
End Property


Public Property Get OrgSuch$()
OrgSuch$ = iOrgSuch$
End Property
Public Property Let OrgSuch(ByVal vNewValue$)
iOrgSuch$ = vNewValue$
End Property


Public Property Get MatchcodePzn$()
MatchcodePzn$ = iMatchcodePzn$
End Property
Public Property Let MatchcodePzn(ByVal vNewValue$)
iMatchcodePzn$ = vNewValue$
End Property


Public Property Get MatchcodeTxt$()
MatchcodeTxt$ = iMatchcodeTxt$
End Property
Public Property Let MatchcodeTxt(ByVal vNewValue$)
iMatchcodeTxt$ = vNewValue$
End Property


Public Property Get MatchcodeErg$()
MatchcodeErg$ = iMatchcodeErg$
End Property
Public Property Let MatchcodeErg(ByVal vNewValue$)
iMatchcodeErg$ = vNewValue$
End Property


Public Property Get UBoundMatchAnzeigeTyp%()
UBoundMatchAnzeigeTyp% = UBound(iMatchAnzeigeTyp$)
End Property
Public Property Let UBoundMatchAnzeigeTyp(ByVal vNewValue%)
'UBound(iMatchAnzeigeTyp$)= vNewValue%
End Property



Public Property Get WarteMaske%()
WarteMaske% = iWarteMaske%
End Property
Public Property Let WarteMaske(ByVal vNewValue%)

If (MusterSuche) Then
    If (vNewValue%) And (iWarteMaske = 0) Then
        Call StartAnimation(frmMatchcode, "Artikel werden gesucht ...")
    ElseIf (vNewValue% = 0) And (iWarteMaske) Then
        Call StopAnimation(frmMatchcode)
    End If
    iWarteMaske% = vNewValue%
    Abbruch = 0
End If

End Property

Public Property Get Abbruch%()
Abbruch% = iAbbruch
End Property
Public Property Let Abbruch(ByVal vNewValue%)
iAbbruch% = vNewValue%
End Property

Sub StartAnimation(hForm As Object, Optional text$ = "Aufgabe wird bearbeitet ...")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StartAnimation")
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

With hForm
    .MousePointer = vbHourglass
    .aniAnimation.Open "findfile.avi"
    .lblAnimation.Caption = text$
    .picAnimationBack.Left = (.ScaleWidth - .picAnimationBack.Width) / 2
    .picAnimationBack.Top = (.ScaleHeight - .picAnimationBack.Height) / 2
'    .picAnimationBack.Left = .Left + (.ScaleWidth - .picAnimationBack.Width) / 2
'    .picAnimationBack.Top = .Top + (.ScaleHeight - .picAnimationBack.Height) / 2
    .picAnimationBack.Visible = True
    .aniAnimation.Play
    .Refresh
End With

Call clsError.DefErrPop
End Sub

Sub StopAnimation(hForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StopAnimation")
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

With hForm
    .aniAnimation.Stop
    .picAnimationBack.Visible = False
    .MousePointer = vbDefault
End With

Call clsError.DefErrPop
End Sub

Sub MakeTaxeSuch(Such$, InhMenge&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeTaxeSuch")
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
'Dim i%, l%, LeftLen%, InhMengeFlag%, InhLen%
'Dim ch$
'
'l% = Len(Such$)
'LeftLen% = l%
'InhMenge& = 0
'
'ch$ = Mid$(Such$, 1, 1)
'If ((ch$ < "0") Or (ch$ > "9")) Then
'    For i% = 1 To l%
'        ch$ = Mid$(Such$, i%, 1)
'        'If ((ch$ >= "0") And (ch$ <= "9")) Then      '2.0.24
'        If ((ch$ >= "0") And (ch$ <= "9")) And (InhLen% <= 7) Then
'            If (InhMengeFlag% = False) Then
'                LeftLen% = i% - 1
'                InhMengeFlag% = True
'            End If
'            InhLen% = InhLen% + 1
'            InhMenge& = (InhMenge& * 10) + (Asc(ch$) - 48)
'        ElseIf (InhMengeFlag% = True) Then
'            Exit For
'        End If
'    Next i%
'End If
'Such$ = Left$(Such$, LeftLen%)

Dim ind%
Do
    ind = InStr(Such, " ")
    If (ind > 0) Then
        Such = Left(Such, ind - 1) + Mid(Such, ind + 1)
    Else
        Exit Do
    End If
Loop

Call clsError.DefErrPop
End Sub


Function MakeTaxeMatchcode(Name As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeTaxeMatchcode")
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
Dim ch$, chAlt$, h$, h2$, zw$, i%, l%, p%, ind%

h$ = RTrim$(Name$)
l% = Len(h$)

For i% = 1 To l%
    ind% = InStr(MatchKlein$, Mid$(h$, i%, 1))
    If (ind% > 0) Then Mid$(h$, i%, 1) = Mid$(MatchGross$, ind%, 1)
Next i%

h2$ = ""
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    ind% = InStr(MatchGueltig$, ch$)
    If (ind% > 0) Then h2$ = h2$ + ch$
Next i%

h$ = h2$

For i% = 0 To (ANZMATCHKONVERTIERUNGEN - 1)
    p% = 2
    Do
        ind% = InStr(p%, h$, MatchPhonAlt(i%))
        If (ind% = 0) Then Exit Do
        If (Mid$(h$, ind% - 1, 1) <> ".") Then
            zw$ = Mid$(h$, ind% + Len(MatchPhonAlt(i%)))
            h$ = Left$(h$, ind% - 1) + MatchPhonNeu(i%) + zw$
        Else
            p% = p% + 1
        End If
    Loop
Next i%

l% = Len(h$)
h2$ = ""
chAlt$ = Chr$(1)
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    If (ch$ <> chAlt$) Then h2$ = h2$ + ch$
    chAlt$ = ch$
Next i%

h$ = Left$(h2$ + Space$(10), 10)

MakeTaxeMatchcode$ = RTrim$(h$)

Call clsError.DefErrPop
End Function


Sub ZerlegeMuster(Muster As String, Packung As Integer, SuchNGr As String, herst As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZerlegeMuster")
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

Dim i As Integer, x As Integer

If Left(Muster, 1) = "," Or Left(Muster, 1) = "*" Then Muster = Mid(Muster, 2)
Packung = 0
SuchNGr = ""
herst = ""
i = InStr(Muster, "#")
x = InStr(Muster, "+")
If i > 0 And x > 0 Then
  If i > x Then
    If UCase(Trim(Mid(Muster, i + 1, 1))) = "N" Then
      SuchNGr = UCase(Trim(Mid(Muster, i + 1)))
    Else
      Packung = Val(Mid(Muster, i + 1))
    End If
    Muster = Left(Muster, i - 1)
    herst = Mid(Muster, x + 1)
    Muster = Left(Muster, x - 1)
  Else
    herst = Mid(Muster, x + 1)
    Muster = Left(Muster, x - 1)
    If UCase(Trim(Mid(Muster, i + 1, 1))) = "N" Then
      SuchNGr = UCase(Trim(Mid(Muster, i + 1)))
    Else
      Packung = Val(Mid(Muster, i + 1))
    End If
    Muster = Left(Muster, i - 1)
  End If
ElseIf i > 0 Then
  If UCase(Trim(Mid(Muster, i + 1, 1))) = "N" Then
    SuchNGr = UCase(Trim(Mid(Muster, i + 1)))
  Else
    Packung = Val(Mid(Muster, i + 1))
  End If
  Muster = Left(Muster, i - 1)
ElseIf x > 0 Then
  herst = Mid(Muster, x + 1)
  Muster = Left(Muster, x - 1)
End If
herst = UCase(RTrim(herst))
Muster = UCase(Muster)

Call clsError.DefErrPop
End Sub

'Function CheckOk%(aus%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("CheckOk%")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim ret%
'Dim CheckText$
'
'ret = True
'If MusterSuche Then
''    MusterSucheBreak = False
''    InMusterSuche = True
''    On Error Resume Next
''    If ActiveMatch < 2 Then
''        If satz > 0 Then
''            MatchTaxeRec.Bookmark = tsatz
''        End If
''    Else
''    '.    GetDataRecord db, satz
''    End If
''
''    ok = True
''    If Err.Number <> 0 Then
''        ok = False
''        On Error GoTo DefErr
''        Return
''    End If
''    On Error GoTo DefErr
'
''    RezPfl = False
''    ApoPfl = False
'    If (iMatchModus% = LAGER_MATCH) Then
'      'war ein StammMatchcode
'        CheckText = Ast1.kurz
'        If Packung > 0 And Val(Ast1.meng) <> Packung Then ret = False
'        If herst > "" And Left(Ast1.herst, Len(herst)) <> herst Then ret = False
''        If InStr(Ast1.rez, "+") > 0 Or InStr(Ast1.rez, "SG") > 0 Then RezPfl = True: ApoPfl = True
'    Else
'        CheckText = UCase(MatchTaxeRec!Leftkomprimiert)
'        If Packung > 0 And Val(MatchTaxeRec!menge) <> Packung Then ret = False
'        If SuchNGr > "" And MatchTaxeRec!NGrösse <> SuchNGr Then ret = False
'        If herst > "" And Left(MatchTaxeRec!HerstellerKB, Len(herst)) <> herst Then ret = False
''        If InStr("0,2,3,4", CStr(MatchTaxeRec!AbgabeBest)) = 0 Then RezPfl = True
''        If MatchTaxeRec!AbgabeBest <> 2 And MatchTaxeRec!AbgabeBest <> 3 Then ApoPfl = True
'        CheckText = UCase(MatchTaxeRec!Name)
'    End If
''    If UCase(Left(CheckText, Len(MusterAnf))) <> MusterAnf Then satz = 0  'Abbruch
''    If ActiveMatch <> 2 Then CheckText = UCase(MatchTaxeRec!Name)
'
''    If ok And (KeineRezPflicht Or KeineApoPflicht) Then
''      If KeineRezPflicht Eqv (Not RezPfl) Then
''        If KeineApoPflicht Eqv (Not ApoPfl) Then
''          If InStr(CheckText, Muster) > 0 Then Checked = True
''        End If
''      End If
''    ElseIf (ok And InStr(CheckText, Muster) > 0) Or satz = 0 Then
''      Checked = True
''    Else
''    End If
'
'    If (ret) Then
'        ret = (InStr(CheckText, Muster) > 0)
'    End If
'    If UCase(Left(CheckText, Len(MusterAnf))) <> MusterAnf Then
'        ret = 0  'Abbruch
'        aus = True
'    End If
'    If (iAbbruch) Then
'        ret = 0  'Abbruch
'        aus = True
'    End If
''    InMusterSuche = False
'    'Satz nur liefern wenn er auch passt
'Else
'  'keine Mustersuche, daher OK
''.  Checked = True
''.  dbAss.ActiveIndex = 1
'End If
'
'CheckOk = ret
'
'Call clsError.DefErrPop
'End Function


