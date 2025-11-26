VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOpTools"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim xFileName$

Dim iSecurPharmFields$()
Dim iSecurPharmValues$()
Dim iSecurPharmNeu%
Dim iSecurPharmDMC As String
Dim iSecurPharmLieferDatum As Integer
Dim iSecurPharmLiefNr As Integer
Dim iSecurPharmBeleg As String
Dim iSecurPharmAbholNr As Integer
            
Private Const DefErrModul = "OPTOOLS.CLS"

Function CVI%(x$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVI%")
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
Dim temp%

CopyMemory temp%, ByVal x$, 2
CVI% = temp%

Call clsError.DefErrPop
End Function

Function CVL&(x$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVL&")
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
Dim temp&

CopyMemory temp&, ByVal x$, 4
CVL& = temp&

Call clsError.DefErrPop
End Function

Function CVS!(x$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVS!")
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
Dim temp!

CopyMemory temp!, ByVal x$, 4
Call DxToIEEEs(temp!)
CVS! = temp!

Call clsError.DefErrPop
End Function

Function CVD#(x$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVD#")
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
Dim temp#

CopyMemory temp#, ByVal x$, 8
Call DxToIEEEd(temp#)
CVD# = temp#

Call clsError.DefErrPop
End Function

Function CVDat%(cDatum$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVDat%")
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
Dim d$

d$ = Right$(cDatum$, 1) + Left$(cDatum$, 1)
CVDat% = CVI%(d$)

Call clsError.DefErrPop
End Function

Function CVDatum$(cDatum$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVDatum$")
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
Dim d$

d$ = Right$(cDatum$, 1) + Left$(cDatum$, 1)
d$ = sDate$(CVI%(d$))
If Len(d$) < 6 Then
    d$ = Space$(8)
Else
    d$ = DateLong$(d$)
End If
CVDatum$ = d$

Call clsError.DefErrPop
End Function

Function OpDatum2String$(ByVal iDatum%, Optional DefaultDatum$ = "01.01.1900")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpDatum2String$")
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
Dim yy%
Dim ret$

If (iDatum < 1) Then
    ret = DefaultDatum$
Else
    ret$ = sDate$(iDatum%)
    yy% = Val(Right$(ret$, 2))
    If yy% >= 50 Then yy% = yy% + 1900 Else yy% = yy% + 2000
    ret = Left(ret, 2) + "." + Mid(ret, 3, 2) + "." + Format(yy, "0000")
End If
OpDatum2String$ = ret

Call clsError.DefErrPop
End Function

Function Date2OpDatum%(ByVal WinDate As Date, Optional DefaultDatum$ = "01.01.1900")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Date2OpDatum%")
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
Dim yy%
Dim ret%

If (Format(WinDate, "DD.MM.YYYY") = DefaultDatum) Then
    ret = 0
ElseIf (Year(WinDate) >= 2050) Then
    ret = 0
Else
    ret = clsOpTool.iDate(Format(WinDate, "DDMMYY"))
End If

Date2OpDatum = ret

Call clsError.DefErrPop
End Function

Function CVDatum2$(ByVal iDatum%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CVDatum2$")
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
Dim d$

d$ = sDate$(iDatum%)
If Len(d$) < 6 Then
    d$ = Space$(8)
Else
    d$ = DateLong$(d$)
End If
CVDatum2$ = d$

Call clsError.DefErrPop
End Function

Function sDate$(iDatum%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("sDate$")
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
Dim i%, xy%, XN%, xm%, xd%
Dim datum$

'* 1.1.1972 bis 31.12.2049 !!

datum$ = ""
If iDatum% >= 1 And iDatum% <= 28490 Then
    xy% = Int((iDatum% - 1) / 365.25) + 1972
    XN% = iDatum% - Int(365.25 * (xy% - 1972) + 0.75)
    xm% = Int((XN% + 31) / 30)
    If Int(30.55 * xm% - 29.95 + 0.001) + 2 * (xm% > 2) - (xm% > 2 And xy% / 4 = Int(xy% / 4)) >= XN% Then xm% = xm% - 1
    xd% = XN% - Int(30.55 * xm% - 29.95 + 0.001) - 2 * (xm% > 2) + ((xm% > 2) And (xy% Mod 4) = 0)
    
    If xy% > 1999 Then xy% = xy% - 2000 Else xy% = xy% - 1900
    datum$ = Right$(Str$(xd%), 2) + Right$(Str$(xm%), 2) + Right$(Str$(xy%), 2)
    For i% = 1 To 6
      If Mid$(datum$, i%, 1) = " " Then Mid$(datum$, i%, 1) = "0"
    Next i%
End If

sDate$ = datum$

Call clsError.DefErrPop
End Function

Function DateLong$(sDatum$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DateLong$")
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
Dim yy%

yy% = Val(Right$(sDatum$, 2))
If yy% >= 50 Then yy% = yy% + 1900 Else yy% = yy% + 2000
DateLong$ = Right$(Str$(yy%), 4) + Mid$(sDatum$, 3, 2) + Left$(sDatum$, 2)

Call clsError.DefErrPop
End Function

Function MKI$(x%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MKI$")
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
Dim temp$

temp$ = String(2, 0)
CopyMemory ByVal temp$, x%, 2
MKI$ = temp$

Call clsError.DefErrPop
End Function

Function MKL$(x&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MKL$")
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
Dim temp$

temp$ = String(4, 0)
CopyMemory ByVal temp$, x&, 4
MKL$ = temp$

Call clsError.DefErrPop
End Function

Function MKS$(x!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MKS$")
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
Dim temp$

temp$ = String(4, 0)
CopyMemory ByVal temp$, x!, 4
MKS$ = temp$

Call clsError.DefErrPop
End Function

Function MKD$(x#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MKD$")
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
Dim temp$

temp$ = String(8, 0)
Call DxToMBFd(x#)
CopyMemory ByVal temp$, x#, 8
MKD$ = temp$

Call clsError.DefErrPop
End Function

Function MKDate$(iDatum%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MKDate$")
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
Dim d$

d$ = MKI$(iDatum%)
MKDate$ = Right$(d$, 1) + Left$(d$, 1)

Call clsError.DefErrPop
End Function

Function iDate%(sDatum$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("iDate%")
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
Dim x%, xd%, xm%, xy%
Dim xdatum$

'* 1.1.1972 (= 1) bis 31.12.2049 (=28490)
x% = 0: xdatum$ = sDatum$
If (Len(xdatum$) = 6) Then
    Do
        x% = InStr(xdatum$, " ")
        If (x% > 0) Then Mid$(xdatum$, x%, 1) = "0"
    Loop While x% > 0
    xd% = Val(Mid$(xdatum$, 1, 2))
    xm% = Val(Mid$(xdatum$, 3, 2))
    xy% = Val(Mid$(xdatum$, 5, 2))
    If (xy% < 50) Then xy% = xy% + 2000 Else xy% = xy% + 1900
    x% = Int(365.25 * (xy% - 1972) + 0.75) + Int(30.55 * xm% - 29.95 + 0.001)
    x% = x% + 2 * (xm% > 2) + xd% - ((xm% > 2) And xy% / 4 = Int(xy% / 4))
    If (xdatum$ <> sDate$(x%)) Then x% = 0
End If
iDate% = x%
Call clsError.DefErrPop
End Function

Function HoleKundenName$(KundenNr)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleKundenName$")
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
Dim ErrNumber%
Dim clsKunden1 As New clsKunden
Dim kDB As Database
Dim kRec As Recordset
Dim kAdoRec As New ADODB.Recordset
Dim h$, SQLStr$
Dim KundenDB1 As clsKundenDB

ErrNumber = 99
If (Para1.Land = "D") Then
    Set KundenDB1 = New clsKundenDB
    KundenDBok = KundenDB1.DBvorhanden
    If (KundenDBok) Then
        KundenDBok = KundenDB1.OpenDB
    End If
    If (KundenDBok) Then
        ErrNumber = 0
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(KundenNr)
        kAdoRec.Open SQLStr$, KundenDB1.ActiveConn
        If (kAdoRec.EOF = False) Then
            h$ = Trim(Trim(clsOpTool.CheckNullStr(kAdoRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kAdoRec!Name)))
        End If
        kAdoRec.Close
        KundenDB1.CloseDB
    Else
        On Error Resume Next
        Err.Clear
        Set kDB = OpenDatabase("kunden.mdb", False, True)
        ErrNumber = Err.Number
        On Error GoTo DefErr
        Err.Clear
        If (ErrNumber = 0) Then
            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr=" + Str$(KundenNr)
            Set kRec = kDB.OpenRecordset(SQLStr$)
            If Not (kRec.EOF) Then
                h$ = Trim(Trim(clsOpTool.CheckNullStr(kRec!VorName)) + " " + Trim(clsOpTool.CheckNullStr(kRec!Name)))
            End If
            kDB.Close
        End If
    End If
End If
If (ErrNumber = 0) Then
Else
    clsKunden1.OpenDatei
    clsKunden1.GetRecord (KundenNr + 1)
    
    h$ = RTrim$(clsKunden1.Wert("name", 0))
    h$ = h$ + " " + RTrim$(clsKunden1.Wert("name", 1))
    
    Call OemToChar(h$, h$)
    
    clsKunden1.CloseDatei
End If

Set clsKunden1 = Nothing

HoleKundenName$ = h$

Call clsError.DefErrPop
End Function

Function Bcd2ascii(ByVal bcd As String, Stellen As Integer) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Bcd2ascii")
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
Dim i As Integer
Dim ascii As String
'ascii = Ergebnis Zahlenstring
'bcd = binärcodierte Zahl
'stellen = Anzahl der Stellen in Ergebnisfeld
'z.B.: stellen = 3   bcd = chr(&H45) + chr(&H67) => ascii = "456"

ascii = String(Stellen, "0")
If bcd <> String(Len(bcd), "*") Then
  For i = 1 To Stellen Step 2
    Mid(ascii, i, 1) = Chr(Int(Asc(Mid(bcd, Int(i / 2) + 1, 1)) / 16) + &H30)
    If i < Stellen Then
      Mid(ascii, i + 1, 1) = Chr((Asc(Mid(bcd, Int(i / 2) + 1, 1)) Mod 16) + &H30)
    End If
  Next i
End If
Bcd2ascii = ascii

Call clsError.DefErrPop
End Function

Function Ascii2bcd(ByVal ascii As String, Stellen As Integer) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Ascii2bcd")
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
Dim i As Integer
Dim bcd As String
Dim bcdwert As Integer

ascii = Right$(String$(Stellen, "0") + ascii, Stellen)
For i = 1 To Stellen
  If InStr("0123456789", Mid$(ascii, i, 1)) = 0 Then
    Mid$(ascii, i, 1) = "0"
  End If
Next i

bcd = String$(Int((Stellen + 1)) / 2, 0)
For i = 1 To Stellen Step 2
  bcdwert = (Asc(Mid(ascii, i, 1)) - &H30) * 16
  If i < Stellen Then
    bcdwert = bcdwert + Asc(Mid$(ascii, i + 1, 1)) - &H30
  End If
  Mid$(bcd$, Int(i / 2) + 1, 1) = Chr$(bcdwert)
Next i

Ascii2bcd = bcd

Call clsError.DefErrPop
End Function

Sub EanPruef(ean As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EanPruef")
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
Dim i As Integer, PruefSumme As Integer, PruefZiffer As Integer

If Val(ean) > 0 And Len(ean) = 13 Then
  PruefSumme = 0
  For i = 1 To 12
    PruefSumme = PruefSumme + Val(Mid(ean, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
  Next i
  PruefZiffer = 10 - (PruefSumme Mod 10)
  Mid(ean, 13, 1) = Right(Str(PruefZiffer), 1)
End If

Call clsError.DefErrPop
End Sub

Sub Taxe2ast(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Taxe2ast")
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
Dim i%, wg%, Appli%, nAm%
Dim dpreis#
Dim bZuz(2) As Byte, by As Byte, by2 As Byte
Dim SQLStr$, h$, h2$, x$, ascii$, gc$, mw$, rez$, arez$
Dim recTaxe As Recordset

If (Para1.Land <> "D") Then
    Call Taxe2astA(pzn$): Call clsError.DefErrPop: Exit Sub
End If

If (TaxeAdoDBok) Then
    Call Taxe2astADO(pzn$): Call clsError.DefErrPop: Exit Sub
End If

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
If (recTaxe.EOF = True) Then Call clsError.DefErrPop: Exit Sub


'Call WGconvEinlesen

'* pzn kurz meng
Ast1.pzn = Format$(recTaxe!pzn, "0000000")
Ast1.kurz = Left$(Left$(recTaxe!Name, 28) + Space$(28), 28)
Ast1.meng = Right$(recTaxe!menge, 5)

'* meh
Ast1.meh = recTaxe!einheit

'* aep
dpreis# = recTaxe!EK / 100
'Call DxToMBFd(dPreis#)
Ast1.aep = dpreis#

'* Festbetrag
dpreis# = recTaxe!FESTBETRAG / 100
'Call DxToMBFd(dPreis#)
Ast1.kp = dpreis#

'* avp
dpreis# = recTaxe!vk / 100
'Call DxToMBFd(dPreis#)
Ast1.AVP = dpreis#

'* mwst
mw$ = "2"
If (recTaxe!MWSTKz) Then mw$ = "1"
Ast1.mw = mw$

'* ka
Ast1.ka = " "

'* wg
'wgconv!!!!??????????
wg% = recTaxe!Warengruppe
Appli% = 0
'If (wg% >= 10) And (wg% <= 19) Then
'  Appli% = -1
'  wg% = wg% - 10
'End If
Ast1.wg = WGconvert(wg%)

'* herst
Ast1.herst = recTaxe!HerstellerKB

'* lic
h$ = recTaxe!UeberGh
h2$ = " "
If (h$ = "1") Then
    h2$ = "D"
ElseIf (h$ = "2") Then
    h2$ = "K"
End If
Ast1.lic = h2$

'* gültig
ascii$ = recTaxe!gueltig
gc$ = Mid$("0123456789OEZ", Val(Mid$(ascii$, 3, 2)) + 1, 1) + Right$(ascii$, 1)

'* vc gc
Ast1.vc = "N"
Ast1.gc = gc$

'* lac
'h$ = Chr$(recTaxe!Lagerung)
'If (Val(h$) = 0) Then h$ = " "
If (recTaxe!Lagerung = 0) Then
    h$ = " "
Else
    h$ = Format(recTaxe!Lagerung, "0")
End If
Ast1.lac = h$


'* rez
'rez$ = Chr$(recTaxe!AbgabeBest)
rez$ = Format$(recTaxe!AbgabeBest, "0")
nAm% = 0: If (InStr("23", rez$) <> 0) Then nAm% = -1
If (Appli%) Then
    arez$ = "A ": If (nAm%) Then arez$ = "AP"
ElseIf (recTaxe!BTMKz) Then
    arez$ = "SG"
ElseIf (InStr("156", rez$) <> 0) Then
    arez$ = "+ "
ElseIf nAm% Then
    'Nicht-Arzneimittel oder nicht apothekenpflichtiges Arzneimittel
    arez$ = "P "
Else
    arez$ = "  "
End If
Ast1.rez = arez$

'* kas
h$ = "   "
If (recTaxe!RufKz) Then
    h$ = "RW "
ElseIf (recTaxe!NegListe > 0) Then
    h$ = "*  "
End If
Ast1.kas = h$

'* kasz
h$ = "  "
If (recTaxe!FestKz > 0) Then
  h$ = "F "
End If
Ast1.kasz = h$

Ast1.zuza = 0    'fehlt noch
h$ = recTaxe!Zuzahlung
If (Left$(h$, 1) = "9") Then Mid$(h$, 1, 1) = "7"

bZuz(0) = 0: bZuz(1) = 0
If (Trim(h$) <> "") Then
    by = Asc(Mid$(h$, 3, 1)) - Asc("0")
    If (by And &H1) Then bZuz(0) = bZuz(0) Or &H1
    If (by And &H2) Then bZuz(0) = bZuz(0) Or &H2
    by2 = Asc(Mid$(h$, 1, 1)) - Asc("0")
    bZuz(0) = bZuz(0) Or (by2 * 32)
    by2 = Asc(Mid$(h$, 2, 1)) - Asc("0")
    bZuz(0) = bZuz(0) Or (by2 * 4)
    
    If (by And &H1) Then bZuz(1) = bZuz(1) Or &H80
    by2 = Asc(Mid$(h$, 4, 1)) - Asc("0")
    bZuz(1) = bZuz(1) Or (by2 * 16)
    by2 = Asc(Mid$(h$, 5, 1)) - Asc("0")
    bZuz(1) = bZuz(1) Or (by2 * 2)
End If
Ast1.zuza = bZuz(0) * 256& + bZuz(1)

Ast1.Vnr = Space$(8)

'* kz
Ast1.kz = " "

'* abl
h$ = " "
If (recTaxe!VerfallArt = "1") Then
  h$ = "A"
End If
Ast1.abl = h$

Ast1.frei = Chr$(0)

Call clsError.DefErrPop
End Sub

Sub Taxe2astADO(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Taxe2astADO")
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
Dim i%, wg%, Appli%, nAm%
Dim dpreis#
Dim bZuz(2) As Byte, by As Byte, by2 As Byte
Dim SQLStr$, h$, h2$, x$, ascii$, gc$, mw$, rez$, arez$
Dim recTaxe As New ADODB.Recordset

'If (Para1.Land <> "D") Then
'    Call Taxe2astA(pzn$): Call clsError.DefErrPop: Exit Sub
'End If
'
'If (TaxeAdoDBok) Then
'    Call Taxe2astADO(pzn$): Call clsError.DefErrPop: Exit Sub
'End If

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
recTaxe.Open SQLStr, TaxeAdoDB1.ActiveConn
If (recTaxe.EOF = True) Then Call clsError.DefErrPop: Exit Sub


'Call WGconvEinlesen

'* pzn kurz meng
Ast1.pzn = clsOpTool.PznString(recTaxe!pzn)
Ast1.kurz = Left$(Left$(recTaxe!Name, 28) + Space$(28), 28)
Ast1.meng = Right$(recTaxe!menge, 5)
   
Ast1.KurzNeu = Trim(recTaxe!Name)
Ast1.MengNeu = Trim(recTaxe!menge)

'* meh
Ast1.meh = recTaxe!einheit

'* aep
dpreis# = recTaxe!EK / 100
'Call DxToMBFd(dPreis#)
Ast1.aep = dpreis#

'* Festbetrag
dpreis# = recTaxe!FESTBETRAG / 100
'Call DxToMBFd(dPreis#)
Ast1.kp = dpreis#

'* avp
dpreis# = recTaxe!vk / 100
'Call DxToMBFd(dPreis#)
Ast1.AVP = dpreis#

'* mwst
mw$ = "2"
If (recTaxe!MWSTKz) Then mw$ = "1"
Ast1.mw = mw$

'* ka
Ast1.ka = " "

'* wg
'wgconv!!!!??????????
wg% = recTaxe!Warengruppe
Appli% = 0
'If (wg% >= 10) And (wg% <= 19) Then
'  Appli% = -1
'  wg% = wg% - 10
'End If
Ast1.wg = WGconvert(wg%)

'* herst
Ast1.herst = recTaxe!HerstellerKB

'* lic
h$ = recTaxe!UeberGh
h2$ = " "
If (h$ = "1") Then
    h2$ = "D"
ElseIf (h$ = "2") Then
    h2$ = "K"
End If
Ast1.lic = h2$

'* gültig
ascii$ = recTaxe!gueltig
gc$ = Mid$("0123456789OEZ", Val(Mid$(ascii$, 3, 2)) + 1, 1) + Right$(ascii$, 1)

'* vc gc
Ast1.vc = "N"
Ast1.gc = gc$

'* lac
'h$ = Chr$(recTaxe!Lagerung)
'If (Val(h$) = 0) Then h$ = " "
If (recTaxe!Lagerung = 0) Then
    h$ = " "
Else
    h$ = Format(recTaxe!Lagerung, "0")
End If
Ast1.lac = h$


'* rez
'rez$ = Chr$(recTaxe!AbgabeBest)
rez$ = Format$(recTaxe!AbgabeBest, "0")
nAm% = 0: If (InStr("23", rez$) <> 0) Then nAm% = -1
If (Appli%) Then
    arez$ = "A ": If (nAm%) Then arez$ = "AP"
ElseIf (recTaxe!BTMKz) Then
    arez$ = "SG"
ElseIf (InStr("156", rez$) <> 0) Then
    arez$ = "+ "
ElseIf nAm% Then
    'Nicht-Arzneimittel oder nicht apothekenpflichtiges Arzneimittel
    arez$ = "P "
Else
    arez$ = "  "
End If
Ast1.rez = arez$

'* kas
h$ = "   "
If (recTaxe!RufKz) Then
    h$ = "RW "
ElseIf (recTaxe!NegListe > 0) Then
    h$ = "*  "
End If
Ast1.kas = h$

'* kasz
h$ = "  "
If (recTaxe!FestKz > 0) Then
  h$ = "F "
End If
Ast1.kasz = h$

Ast1.zuza = 0    'fehlt noch
h$ = recTaxe!Zuzahlung
If (Left$(h$, 1) = "9") Then Mid$(h$, 1, 1) = "7"

bZuz(0) = 0: bZuz(1) = 0
If (Trim(h$) <> "") Then
    by = Asc(Mid$(h$, 3, 1)) - Asc("0")
    If (by And &H1) Then bZuz(0) = bZuz(0) Or &H1
    If (by And &H2) Then bZuz(0) = bZuz(0) Or &H2
    by2 = Asc(Mid$(h$, 1, 1)) - Asc("0")
    bZuz(0) = bZuz(0) Or (by2 * 32)
    by2 = Asc(Mid$(h$, 2, 1)) - Asc("0")
    bZuz(0) = bZuz(0) Or (by2 * 4)
    
    If (by And &H1) Then bZuz(1) = bZuz(1) Or &H80
    by2 = Asc(Mid$(h$, 4, 1)) - Asc("0")
    bZuz(1) = bZuz(1) Or (by2 * 16)
    by2 = Asc(Mid$(h$, 5, 1)) - Asc("0")
    bZuz(1) = bZuz(1) Or (by2 * 2)
End If
Ast1.zuza = bZuz(0) * 256& + bZuz(1)

Ast1.Vnr = Space$(8)

'* kz
Ast1.kz = " "

'* abl
h$ = " "
If (recTaxe!VerfallArt = "1") Then
  h$ = "A"
End If
Ast1.abl = h$

Ast1.frei = Chr$(0)

recTaxe.Close

Call clsError.DefErrPop
End Sub

Sub Taxe2astA(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Taxe2astA")
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
Dim i%, wg%, Appli%, nAm%
Dim dpreis#
Dim bZuz(2) As Byte, by As Byte, by2 As Byte
Dim SQLStr$, h$, h2$, x$, ascii$, gc$, mw$, rez$, arez$
Dim recTaxe As Recordset

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
If (recTaxe.EOF = True) Then Call clsError.DefErrPop: Exit Sub


'Call WGconvEinlesen

'* pzn kurz meng
Ast1.pzn = Format$(recTaxe!pzn, "0000000")
Ast1.kurz = Left$(Left$(recTaxe!Name, 28) + Space$(28), 28)
Ast1.meng = Right$(recTaxe!menge, 5)

'* meh
Ast1.meh = recTaxe!einheit

'* aep
dpreis# = recTaxe!EK / 100
'Call DxToMBFd(dPreis#)
Ast1.aep = dpreis#

'* Festbetrag
dpreis# = recTaxe!FESTBETRAG / 100
'Call DxToMBFd(dPreis#)
Ast1.kp = dpreis#

'* avp
dpreis# = recTaxe!vk / 100
'Call DxToMBFd(dPreis#)
Ast1.AVP = dpreis#

'* mwst
Ast1.mw = Format(recTaxe!MWSTKz, "0")

'* ka
Ast1.ka = recTaxe!kzaufschla

'* wg
Ast1.wg = Mid$(Str$(recTaxe!Warengruppe) + "  ", 2, 2)

'* herst
Ast1.herst = recTaxe!HerstellerKB

'* lic
Ast1.lic = recTaxe!lica

'* vc
Ast1.vc = recTaxe!ArtStatus
'* gc
ascii$ = recTaxe!gueltig
gc$ = Mid$("0123456789OEZ", Val(Mid$(ascii$, 3, 2)) + 1, 1) + Right$(ascii$, 1)
Ast1.gc = gc$

'* lac
Ast1.lac = recTaxe!LagerungA

'* rez
Ast1.rez = recTaxe!RezA

'* kas
Ast1.kas = recTaxe!DarForm


'* kasz
Ast1.kasz = recTaxe!KasZuA

Ast1.zuza = 0    'fehlt noch
Ast1.Vnr = Right$(recTaxe!vnra, 8)

'* kz
Ast1.kz = Chr$(recTaxe!verfallmonate)

'* abl
Ast1.abl = Chr$(recTaxe!VerfallArt)

Ast1.frei = Chr$(0)

Call clsError.DefErrPop
End Sub
    
Function WGconvert$(wg%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WGconvert$")
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
Dim NeuWg%, iw%
Dim TestWg$, x$

'NeuWg% = 0: TestWg$ = "," + Mid$(Str$(wg%), 2) + ","
'For iw% = 1 To 9
'  If (NeuWg% = 0) Then
' x$ = Para1.WGconv(iw%): If (InStr(x$, TestWg$) <> 0) Then NeuWg% = iw%
'  End If
'Next iw%
'If (NeuWg% = 0) Then NeuWg% = 9

'WGconvert$ = Right$(Str$(NeuWg%), 1) + " "

NeuWg% = wg%
WGconvert$ = Left$(Format(NeuWg%, "0") + Space$(2), 2)

Call clsError.DefErrPop
End Function

Sub WarteAufTaskEnde(SuchTaskId&, pForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WarteAufTaskEnde")
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
Dim erg%

pForm.Enabled = False
Do
    erg% = CheckTask%(SuchTaskId&, pForm)
    If (erg% = False) Then Exit Do
    DoEvents
Loop
pForm.Enabled = True

Call clsError.DefErrPop
End Sub

Function CheckTask%(SuchTaskId&, pForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckTask%")
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
Dim ThreadID
Dim CurrWnd&, x&

CheckTask% = True

CurrWnd& = GetWindow(pForm.hWnd, GW_HWNDFIRST)
While (CurrWnd& <> 0)
  x& = GetWindowThreadProcessId(CurrWnd&, ThreadID)
  If (ThreadID = SuchTaskId&) Then Call clsError.DefErrPop: Exit Function
  CurrWnd& = GetWindow(CurrWnd&, GW_HWNDNEXT)
Wend

CheckTask% = False
Call clsError.DefErrPop
End Function

Function MachStammAuswahlZuzahlung$(zuza%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachStammAuswahlZuzahlung$")
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
Dim i%, z%, z1%, z2%, zz%
Dim ZuzWert!, zu!
Dim ch$, Zuzahl$

Zuzahl$ = ""
zu! = zuza%
If (zu! < 0) Then zu! = zu! + 65536!
'* 1 x shift right
z% = (zu! / 2)
For z1% = 1 To 5
  z2% = (z% Mod 8)
  '* 3 x shift right
  z% = z% \ 8
  Zuzahl$ = Right$(Str$(z2%), 1) + Zuzahl$
Next z1%
        

ch$ = Left$(Zuzahl$, 1)
'If (tRec!Warengruppe = 14) Then ch$ = "1"
If (ch$ = "9") Then
    MachStammAuswahlZuzahlung$ = "20%"
ElseIf (ch$ = "0") Then
    MachStammAuswahlZuzahlung$ = "kA"
ElseIf (ch$ = "1") Then
    MachStammAuswahlZuzahlung$ = "nb"
ElseIf (ch$ = "2") Then
    MachStammAuswahlZuzahlung$ = "ng"
Else
    ZuzWert! = 0
    For i% = 1 To 5
        zz% = Val(Mid$(Zuzahl$, i%, 1))
        ZuzWert! = ZuzWert! + Para1.ZuzWert(zz%)
    Next i%
        
    MachStammAuswahlZuzahlung$ = Format$(ZuzWert!, "#0.00")
End If

Call clsError.DefErrPop
End Function

Function FNX#(Wert#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FNX#")
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

FNX# = Sgn(Wert#) * CDbl(Int(Abs(Wert#) * 100# + 0.501) / 100#)

Call clsError.DefErrPop
End Function

'Function ModemParameter%(TestName$, xFilePara$, Optional Parameter% = False)
Function ModemParameter%(TestName$, xFilePara$, Parameter%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ModemParameter%")
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
Dim i%, ok%, xq%, xParam%, DPA%, aa%
Dim s$, x$, xGeraet$, xParams$(6)
Dim XparamFix As String * 128
Dim DparamFix As String * 40

x$ = Para1.user
If (Val(x$) = 0) Then x$ = ""
s$ = "\user\xparam" + x$ + ".dat"
xParam% = clsDat.FileOpen(s$, "RW", "R", Len(XparamFix))

'1 * belegt sonst nicht belegt, 8 Name,30 Schnittstelle incl.params
i% = 1: ok% = False
xFilePara$ = ""
While (i% <= 9) And (ok% = 0)
    Get xParam%, i% + 1, XparamFix
    If (Left$(XparamFix$, 1) = "*") Then
        x$ = Mid$(XparamFix$, 2, 8): x$ = LTrim$(x$): x$ = RTrim$(x$): xGeraet$ = x$
        xFileName$ = "\user\" + x$ + ".dpa"
        x$ = Mid$(XparamFix$, 10, 30): x$ = LTrim$(x$): x$ = RTrim$(x$): xFilePara$ = x$
        If (Left$(xGeraet$, Len(TestName$)) = TestName$) Then ok% = True
    End If
    i% = i% + 1
Wend
Close #xParam%

If (ok% And Parameter%) Then
    DPA% = clsDat.FileOpen(xFileName$, "RW", "R", Len(DparamFix))
    For xq% = 1 To 5
        Get #DPA%, xq%, DparamFix: x$ = DparamFix
        x$ = RTrim$(x$): xParams$(xq%) = x$
    Next xq%
    Close #DPA%
End If

x$ = Mid$(xParams$(5), 21): x$ = LTrim$(x$): x$ = RTrim$(x$): xParams$(6) = x$
x$ = Left$(xParams$(5), 20): x$ = LTrim$(x$): x$ = RTrim$(x$): xParams$(5) = x$

ModemParameter% = ok%

Call clsError.DefErrPop
End Function

Sub GetModemBefehle(xParams$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetModemBefehle")
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
Dim xq%, DPA%
Dim x$
Dim DparamFix As String * 40

DPA% = clsDat.FileOpen(xFileName$, "RW", "R", Len(DparamFix))
For xq% = 1 To 5
    Get #DPA%, xq%, DparamFix: x$ = DparamFix
    x$ = RTrim$(x$): xParams$(xq%) = x$
Next xq%
Close #DPA%

x$ = Mid$(xParams$(5), 21): x$ = LTrim$(x$): x$ = RTrim$(x$): xParams$(6) = x$
x$ = Left$(xParams$(5), 20): x$ = LTrim$(x$): x$ = RTrim$(x$): xParams$(5) = x$

Call clsError.DefErrPop
End Sub

Function GetGeraetParameter%(ind%, GeraetName$, GeraetPara$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetGeraetParameter%")
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
Dim ok%, xParam%
Dim x$, s$
Dim XparamFix As String * 128

x$ = Para1.user
If (Val(x$) = 0) Then x$ = ""
s$ = "\user\xparam" + x$ + ".dat"
xParam% = clsDat.FileOpen(s$, "RW", "R", Len(XparamFix))

'1 * belegt sonst nicht belegt, 8 Name,30 Schnittstelle incl.params
ok% = False
GeraetName$ = ""
GeraetPara$ = ""

Get xParam%, ind% + 1, XparamFix
If (Left$(XparamFix$, 1) = "*") Then
    x$ = Mid$(XparamFix$, 2, 8): x$ = LTrim$(x$): x$ = RTrim$(x$): GeraetName$ = x$
    x$ = Mid$(XparamFix$, 10, 30): x$ = LTrim$(x$): x$ = RTrim$(x$): GeraetPara$ = x$
    ok% = True
End If

Close #xParam%

GetGeraetParameter% = ok%

Call clsError.DefErrPop
End Function

Function OpenCom%(sForm As Object, xFilePara$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenCom%")
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
Dim i%, ret%, fehler%, CommPort%, ind%, ind2%, stByte%
Dim h$, Settings$

ret% = True

fehler% = 0
' COM einsetzen.
ind% = InStr(xFilePara$, "COM")
If (ind% > 0) Then
    CommPort% = Val(Mid$(xFilePara$, ind% + 3, 1))
    Settings$ = Mid$(xFilePara$, ind% + 5)
    
    stByte% = 1
    For i% = 1 To 4
        ind2% = InStr(stByte%, Settings$, ",")
        If (ind2% > 0) Then
            If (i% = 4) Then
                Settings$ = Left$(Settings$, ind2% - 1)
            Else
                stByte% = ind2% + 1
            End If
        Else
            Exit For
        End If
    Next i%

End If

With sForm
    .comSenden.CommPort = CommPort%
    .comSenden.Settings = Settings$
    .comSenden.InputMode = comInputModeText
    .comSenden.Handshaking = comRTSXOnXOff
    .comSenden.InputLen = 1

    ' _Anschluß öffnen.
    On Error GoTo ErrorHandler
    .comSenden.PortOpen = True
End With

If (fehler%) Then ret% = False

OpenCom% = ret%

Call clsError.DefErrPop: Exit Function

ErrorHandler:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call clsError.DefErrPop
End Function

Function PruefeDp04Zeile$(Dp04Str$, Dp04Menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PruefeDp04Zeile$")
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
Dim i%, CodeTyp%
Dim PruefSumme&
Dim PruefZiffer As Byte
Dim h$

h$ = Mid$(Dp04Str$, 3)

Dp04Menge% = Val(Mid$(h$, 14, 4))
'If (Dp04Menge% = 0) Then Dp04Menge% = 1

CodeTyp% = Asc(Mid$(h$, 19, 1))
If (CodeTyp% = 1) Then
    '13-stelliger
    h$ = Mid$(h$, 2, 12)
    PruefSumme = 0
    For i% = 1 To 12
        PruefSumme = PruefSumme + Val(Mid$(h$, i%, 1)) * (1 + 2 * ((i% + 1) Mod 2))
    Next i%
    PruefZiffer = 10 - (PruefSumme Mod 10)
    h$ = h$ + Right$(Str$(PruefZiffer), 1)
ElseIf (CodeTyp% = 4) Then
    '8-stelliger EAN
    h$ = Mid$(h$, 6, 8)
Else
    'Kärtchen Typ=2
    h$ = Mid$(h$, 6, 7)
End If

PruefeDp04Zeile$ = h$

Call clsError.DefErrPop
End Function

Function CalcDirektBM%(iBevorratungsZeit%, bmo!, ls%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CalcDirektBM%")
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
Dim ret%, br%, BESUCH%
Dim bv!

ret% = 0

'BvErrechnen:
br% = Para1.BestellPeriode
BESUCH% = iBevorratungsZeit%

bv! = (bmo! / br%) * BESUCH%
If (bv! < 0) Then bv! = 0
If (bv! > ls% + bmo!) Then
    bv! = bv! - ls%
Else
    bv! = bmo!
End If
If br% * ls% >= BESUCH% * (bmo! * 4 / 5) Then bv! = 0
If ls% = 0 And bv! = 0 Then bv! = 1
If bv! > 32000 Then bv! = 32000
If bv! < -32000 Then bv! = -32000

ret% = Int(bv! + 0.501)

CalcDirektBM% = ret%

Call clsError.DefErrPop
End Function


Function xVal#(ByVal x As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("xVal#")
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
Dim i As Integer
Dim j As Integer

i = InStr(x, ",")
j = InStr(x, ".")

If i > 0 Then
  If j = 0 Then
    Mid(x, i, 1) = "."
  Else
    'erst die Tausendertrennzeichen weg
    While j > 0
      x = Left(x, j - 1) + Mid(x, j + 1)
      j = InStr(x, ".")
    Wend
    i = InStr(x, ",")
    Mid(x, i, 1) = "."
  End If
End If
xVal = Val(x)

Call clsError.DefErrPop
End Function

Function MakePzn$(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakePzn$")
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
Dim i%, PruefZiffer%
Dim iw&
Dim ret$

ret$ = ""
If (Len(pzn$) = 8) Then
    iw& = 0
    For i% = 1 To 7
        iw& = iw& + Val(Mid$(pzn$, i%, 1)) * (i%)
    Next i%
    PruefZiffer% = (iw& Mod 11)
    
    ret$ = Left$(pzn$, 7) + Right$(Str$(PruefZiffer%), 1)
End If

MakePzn$ = ret$

Call clsError.DefErrPop
End Function

Function CheckNullByte(s As Variant) As Byte
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckNullByte")
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
Dim ret As Byte

If (IsNull(s)) Then
    ret = 0
Else
    ret = s
End If

CheckNullByte = ret

Call clsError.DefErrPop
End Function

Function CheckNullLong&(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckNullLong&")
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
Dim ret&

If (IsNull(s)) Then
    ret& = 0
Else
    ret& = s
End If

CheckNullLong& = ret&

Call clsError.DefErrPop
End Function

Function CheckNullDouble#(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckNullDouble#")
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
Dim ret#

If (IsNull(s)) Then
    ret# = 0
Else
    ret# = s
End If

CheckNullDouble# = ret#

Call clsError.DefErrPop
End Function

Function CheckNullStr$(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckNullStr$")
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
Dim ret$

If (IsNull(s)) Then
    ret$ = ""
Else
    ret$ = s
End If

CheckNullStr$ = ret$

Call clsError.DefErrPop
End Function

Function CheckNullDate(s As Variant) As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckNullDate")
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
Dim ret As Date

If (IsNull(s)) Then
    ret = "01.01.1980"
Else
    ret = s
End If

CheckNullDate = ret

Call clsError.DefErrPop
End Function

Function ErhoeheCounter&()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ErhoeheCounter&")
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
Dim iCounter&

iCounter& = GetParameter("COUNTER")
iCounter& = (iCounter& + 1) Mod 100
Call PutParameter%("COUNTER", iCounter&)

ErhoeheCounter& = iCounter&

Call clsError.DefErrPop
End Function

Function GetParameter&(ParaName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetParameter&")
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
Dim ret&, lRecs&

ret& = 0
If (AbholerSQL) Then
    On Error Resume Next
    AbholerParaAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    SQLStr = "SELECT * FROM Parameter WHERE Name='" + ParaName + "'"
    With AbholerParaAdoRec
        .Open SQLStr, AbholerConn
        If (.EOF) Then
            SQLStr = "INSERT INTO Parameter (Name) VALUES ('" + ParaName + "')"
            Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
'            .AddNew
'            ParameterAdoRec!Name = ParaName$
'            ParameterAdoRec!wert = 0
'            .Update
        Else
            ret& = CheckNullLong(AbholerParaAdoRec!Wert)
        End If
    End With
Else
    With AbholerParaRec
        .Seek "=", ParaName$
        If (.NoMatch) Then
            .AddNew
            AbholerParaRec!Name = ParaName$
            AbholerParaRec!Wert = 0
            .Update
        Else
            ret& = clsOpTool.CheckNullLong(AbholerParaRec!Wert)
        End If
    End With
End If

GetParameter& = ret&

Call clsError.DefErrPop
End Function

Sub PutParameter(ParaName$, ParaWert&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutParameter")
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
DefErr2:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul, 9999)
Case vbRetry
  Resume
End Select
Call clsError.DefErrPop: Exit Sub
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim lRecs&

On Error GoTo DefErr2

If (AbholerSQL) Then
    On Error Resume Next
    AbholerParaAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    SQLStr = "SELECT * FROM Parameter WHERE Name='" + ParaName + "'"
    With AbholerParaAdoRec
        .Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
        If (.EOF) Then
            SQLStr = "INSERT INTO Parameter (Name) VALUES ('" + ParaName + "')"
            Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
            AbholerParaAdoRec.Close
            SQLStr = "SELECT * FROM Parameter WHERE Name='" + ParaName + "'"
            .Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
        End If
        AbholerParaAdoRec!Wert = ParaWert&
        .Update
    End With
Else
    With AbholerParaRec
        .Seek "=", ParaName$
        If (.NoMatch) Then
            .AddNew
            AbholerParaRec!Name = ParaName$
        Else
            .Edit
        End If
        AbholerParaRec!Wert = ParaWert&
        .Update
    End With
End If

Call clsError.DefErrPop
End Sub

Function SqlString(ByVal s As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlString " + s)
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
Dim i%

SqlString = s

i = InStr(s, Chr(34))
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, Chr(34))
Loop

i = InStr(s, Chr(0))
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, Chr(0))
Loop

i = InStr(s, "'")
Do While i > 0
    If (i < Len(s)) Then
        If (Mid$(s, i + 1, 1) <> "'") Then
            s = Left(s, i) + "'" + Mid(s, i + 1)
            i = i + 1
        End If
        i = InStr(i + 1, s, "'")
    Else
        s = Left(s, i) + "'"
        Exit Do
    End If
Loop

i = InStr(s, "_")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "_")
Loop

i = InStr(s, "%")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "%")
Loop

i = InStr(s, "^")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "^")
Loop

i = InStr(s, "|")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "|")
Loop

i = InStr(s, "[")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "[")
Loop

i = InStr(s, "]")
Do While i > 0
  s = Left(s, i - 1) + " " + Mid(s, i + 1)
  i = InStr(s, "]")
Loop

SqlString = s

Call clsError.DefErrPop
End Function

Function SqlPzn(ByVal s As String, Optional IstZahl% = True) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlPzn " + s)
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
Dim ret$

ret = CStr(xVal(s))
If (IstZahl) Then
Else
    ret = Right$(String(7, "0") + ret, 7)
End If
SqlPzn = ret$

Call clsError.DefErrPop
End Function

Function uFormat(ByVal Wert As Double, ByVal maske As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("uFormat")
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
Dim s As String
Dim i As Integer
Dim Stellen As Integer

Stellen = Len(maske)
s = Format(Abs(Wert), maske)
s = LTrim(s)
If Wert < 0 Then s = "-" + s
''GS: LTRIM genügt nicht bei negativen Zahlen, da minus ganz vorne steht --> Leerzeichen fressen
'i = InStr(s, " ")
'While Len(s) > Stellen And i > 0
'  s = Left(s, i - 1) + Mid(s, i + 1)
'  i = InStr(s, " ")
'Wend
If Len(s) < Stellen Then s = Right(Space(Stellen) + s, Stellen)
i = InStr(s, ","): If i > 0 Then Mid(s, i, 1) = "."
uFormat = s

Call clsError.DefErrPop
End Function

Function PznString$(lPzn&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PznString$")
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

PznString = Right(String(8, "0") + CStr(lPzn), 8)

Call clsError.DefErrPop
End Function


Public Property Get SecurPharmField$(ind%)
SecurPharmField = IIf((ind >= 0) And (ind <= UBound(iSecurPharmFields)), iSecurPharmFields(ind), "")
End Property

Public Property Get SecurPharmValue$(ind%)
SecurPharmValue = IIf((ind >= 0) And (ind <= UBound(iSecurPharmValues)), iSecurPharmValues(ind), "")
End Property

Public Property Get SecurPharmNeu%()
SecurPharmNeu = iSecurPharmNeu
End Property

Public Property Get SecurPharmDMC$()
SecurPharmDMC = iSecurPharmDMC
End Property
Public Property Let SecurPharmDMC(ByVal vNewValue$)
iSecurPharmDMC = vNewValue$
End Property

Public Property Get SecurPharmLieferDatum%()
SecurPharmLieferDatum = iSecurPharmLieferDatum
End Property
Public Property Get SecurPharmLiefNr%()
SecurPharmLiefNr = iSecurPharmLiefNr
End Property
Public Property Get SecurPharmBeleg$()
SecurPharmBeleg = iSecurPharmBeleg
End Property
Public Property Get SecurPharmAbholNr%()
SecurPharmAbholNr = iSecurPharmAbholNr
End Property

'Function CheckSecurPharm$(ScanStr$, Optional ActBeleg$ = "", Optional ActBelegDatum% = 0, Optional ActBelegLief% = 0, Optional AbholNr% = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("CheckSecurPharm$")
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
'Dim i, ind, l, iOk, gef As Integer
'Dim h$, hh$, iScanStr$, ret$
'
'Dim MessageHeader$, MessageTrailer$, FormatHeader$, FormatTrailer$, FieldSeperator$, RecordSeperator$, sField$, sValue$
'Dim ASC_06Macro_Header$, ASC_06Macro_Trailer$
'Dim ASC_DataIdentifier$()
'
'Dim GS1_Header$
'Dim GS1_DataIdentifier$()
'
'Dim LieferDatum As String
'LieferDatum = ""
'If (ActBelegDatum > 0) Then
'    h = sDate(ActBelegDatum)
'    LieferDatum = Mid(h, 1, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5, 2)
'End If
'
''IFA_DataIdentifier = Array("9N", "1T", "S", "D", "16D", "8P")
'ASC_DataIdentifier = Split("9N|1T|S|D|16D|8P", "|")
'GS1_DataIdentifier = Split("01|10|21|17|16D|01", "|")
'
'iSecurPharmFields = Split("Pzn|Charge|SerienNr|Verfalldatum|Herstelldatum|GTIN", "|")
'iSecurPharmValues = Split("|||||", "|")
'iSecurPharmNeu = 0
'iSecurPharmDMC = ""
'iSecurPharmLieferDatum = 0
'iSecurPharmLiefNr = 0
'iSecurPharmBeleg = ""
'iSecurPharmAbholNr = 0
'
'
'FieldSeperator = Chr(29)    'GS
'RecordSeperator = Chr(30)   'RS
'MessageTrailer = Chr(4) 'EOT
'
'FieldSeperator = "|"    'GS
'RecordSeperator = "{"  'RS
'MessageTrailer = "}" 'EOT
'
'h$ = Space(5)
'l = GetPrivateProfileString("Codes", "GS", "|", h$, 6, CurDir + "\SecurPharm.ini")
'FieldSeperator = Left$(h$, l)
'
'h$ = Space(5)
'l = GetPrivateProfileString("Codes", "RS", "{", h$, 6, CurDir + "\SecurPharm.ini")
'RecordSeperator = Left$(h$, l)
'
'h$ = Space(5)
'l = GetPrivateProfileString("Codes", "EOT", "}", h$, 6, CurDir + "\SecurPharm.ini")
'MessageTrailer = Left$(h$, l)
'
'
'
'
'MessageHeader = "[)>" + RecordSeperator
'FormatHeader = "06" + FieldSeperator
'FormatTrailer = RecordSeperator
'
'ASC_06Macro_Header = MessageHeader + FormatHeader
'ASC_06Macro_Trailer = FormatTrailer + MessageTrailer
'
'GS1_Header = GS1_DataIdentifier(0)
'
''For i = 0 To UBound(SecurPharmFields)
''    SecurPharmFields(i) = ""
''Next i
'
'ScanStr$ = Trim(ScanStr)
'
'iScanStr = ScanStr
'ret = iScanStr
'
'l = Len(iScanStr)
'
'If (l > 20) Then
''    If (Left(iScanStr, Len(ASC_06Macro_Header)) = ASC_06Macro_Header) Then
''        iScanStr = Mid(iScanStr, Len(ASC_06Macro_Header) + 1)
''        If (Right(iScanStr, Len(ASC_06Macro_Trailer)) = ASC_06Macro_Trailer) Then
''            iScanStr = Left(iScanStr, Len(iScanStr) - (Len(ASC_06Macro_Trailer)))
''            If (Right(iScanStr, 1) <> FieldSeperator) Then
''                iScanStr = iScanStr + FieldSeperator
''            End If
''
''            Do
''                ind = InStr(iScanStr, FieldSeperator)
''                If (ind <= 0) Then
''                    Exit Do
''                End If
''
''                sField = Left(iScanStr, ind - 1)
''                iScanStr = Mid(iScanStr, ind + 1)
''
''                For i = 0 To UBound(ASC_DataIdentifier)
''                    l = Len(ASC_DataIdentifier(i))
''                    If (Left(sField, l) = ASC_DataIdentifier(i)) Then
''                        sValue = Mid(sField, l + 1)
''                        If (i = 0) Then
''                            If (Len(sValue) = 12) Then
''                                sValue = Mid(sValue, 3, 8)
''                            Else
''                                sValue = ""
''                            End If
''                        ElseIf (i = 3) Then
'''                            sValue = IIf(Mid(sValue, 5, 2) = "00", "01", Mid(sValue, 5, 2)) + Mid(sValue, 3, 2) + Mid(sValue, 1, 2)
''                            If (Mid(sValue, 5, 2) = "00") Then
''                                Mid(sValue, 5, 2) = "01"
''                            End If
''                        End If
''                        iSecurPharmValues(i) = sValue
''                        Exit For
''                    End If
''                Next i
''            Loop
''            ret = iSecurPharmValues(0)
''        End If
''    ElseIf (Left(iScanStr, Len(GS1_Header)) = GS1_Header) Then
''        Do
''            If (iScanStr = "") Then
''                Exit Do
''            End If
''
''            gef = -1
''            For i = 0 To UBound(GS1_DataIdentifier)
''                If (Left(iScanStr, 2) = GS1_DataIdentifier(i)) Then
''                    If (i = 0) Then
''                        sValue = Mid(iScanStr, 3, 14)
''                        iScanStr = Mid(iScanStr, 17)
''
''                        If (Left(sValue, 5) = "04150") Then
''                            sValue = Mid(sValue, 6, 8)
''                        Else
''                            sValue = ""
''                        End If
''                    ElseIf (i = 3) Then
''                        sValue = Mid(iScanStr, 3, 6)
''                        iScanStr = Mid(iScanStr, 9)
''
'''                        sValue = IIf(Mid(sValue, 5, 2) = "00", "01", Mid(sValue, 5, 2)) + Mid(sValue, 3, 2) + Mid(sValue, 1, 2)
''                        If (Mid(sValue, 5, 2) = "00") Then
''                            Mid(sValue, 5, 2) = "01"
''                        End If
''                    Else
''                        ind = InStr(iScanStr, FieldSeperator)
''                        If (ind > 0) Then
''                            sValue = Mid(iScanStr, 3, ind - 3)
''                            iScanStr = Mid(iScanStr, ind + 1)
''                        Else
''                            sValue = Mid(iScanStr, 3)
''                            iScanStr = ""
''                        End If
''                    End If
''                    iSecurPharmValues(i) = sValue
''                    gef = i
''                    Exit For
''                End If
''            Next i
''            If (gef < 0) Then
''                Exit Do
''            End If
''        Loop
''        ret = iSecurPharmValues(0)
''    End If
'
'    Dim IstAsc As Boolean
'    IstAsc = False
'
'    i = InStr(UCase(Left(iScanStr, 10)), "9N")
'    If i = 0 Then   'GS1
'        i = InStr(iScanStr, "0104150")
'        If i = 0 Then
'            If Left(iScanStr, 2) = "01" Then
'                'ausländischer GTIN
'                iScanStr = Mid(iScanStr, 17)
'            Else
'                CheckSecurPharm = ret
'                Call clsError.DefErrPop: Exit Function
'            End If
'        End If
'        If i > 0 And i < 7 Then
'            iScanStr = Mid(iScanStr, i + 7)   '014150 weg
'            iSecurPharmValues(0) = Left(iScanStr, 8)
'            iScanStr = Mid(iScanStr, 10)    '1 PZ
'        ElseIf i > 0 Then
'            iSecurPharmValues(0) = Mid(iScanStr, i + 7, 8)
'            iScanStr = Left(iScanStr, i - 1) + Mid(iScanStr, i + 16)
'        End If
'    Else    'ASC
'        IstAsc = True
'                If (Right(iScanStr, Len(ASC_06Macro_Trailer)) = ASC_06Macro_Trailer) Then
'            iScanStr = Left(iScanStr, Len(iScanStr) - (Len(ASC_06Macro_Trailer)))
'        End If
'        iScanStr = Mid(iScanStr, i + 4)   '9N11 weg
'        iSecurPharmValues(0) = Left(iScanStr, 8)
'        iScanStr = Mid(iScanStr, 11)    '2PZ
'    End If
'
''    ret = 1
'    Do While Len(iScanStr) > 1
'        iScanStr = Trim(iScanStr)
'        If Left(iScanStr, 1) = "|" Or Left(iScanStr, 1) = "<" Then iScanStr = Trim(Mid(iScanStr, 2))
'        If IstAsc Then
'            If UCase(Left(iScanStr, 2)) = "1T" Then    'Charge
'                iScanStr = Mid(iScanStr, 3)
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = InStr(iScanStr, "<")
'                If i = 0 Then i = Len(iScanStr) + 1
'                iSecurPharmValues(1) = Trim(Left(iScanStr, i - 1))
'                iScanStr = Mid(iScanStr, i + 1)
'            ElseIf UCase(Left(iScanStr, 1)) = "S" Then   'Seriennummer
'                iScanStr = Mid(iScanStr, 2)
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = InStr(iScanStr, "<")
'                If i = 0 Then i = Len(iScanStr) + 1
'                iSecurPharmValues(2) = Trim(Left(iScanStr, i - 1))
'                iScanStr = Mid(iScanStr, i + 1)
'            ElseIf UCase(Left(iScanStr, 1)) = "D" Then   'Verfall
'                iSecurPharmValues(3) = Mid(iScanStr, 2, 6)
'                iScanStr = Mid(iScanStr, 8)
'            Else
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = InStr(iScanStr, "<")
'                If i = 0 Then i = Len(iScanStr) + 1
'                If i > 0 Then
'                    iScanStr = Mid(iScanStr, i + 1)
'                End If
'            End If
'        Else
'            Select Case Left(iScanStr, 2)
'            Case "21"   'Seriennummer
'                iScanStr = Mid(iScanStr, 3)
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = Len(iScanStr) + 1
'                iSecurPharmValues(2) = Left(iScanStr, i - 1)
'                iScanStr = Mid(iScanStr, i + 1)
'            Case "10"   'Charge
'                iScanStr = Mid(iScanStr, 3)
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = Len(iScanStr) + 1
'                iSecurPharmValues(1) = Left(iScanStr, i - 1)
'                iScanStr = Mid(iScanStr, i + 1)
'            Case "17"   'Verfall
'                iSecurPharmValues(3) = Mid(iScanStr, 3, 6)
'                iScanStr = Mid(iScanStr, 9)
'            Case "71"   'PZN bei ausländischem GTIN
'                If Val(iSecurPharmValues(0)) = 0 And Left(iScanStr, 3) = "710" Then
'                    iSecurPharmValues(0) = Mid(iScanStr, 4, 8)
'                    iScanStr = Mid(iScanStr, 12)
'                    If Left(iScanStr, 1) = "|" Then iScanStr = Mid(iScanStr, 2)
'                Else
'                    i = InStr(iScanStr, "|")
'                    If i = 0 Then i = Len(iScanStr) + 1
'                    If i > 0 Then iScanStr = Mid(iScanStr, i + 1)
'                End If
'            Case Else
'                i = InStr(iScanStr, "|")
'                If i = 0 Then i = Len(iScanStr) + 1
'                If i > 0 Then
'                    iScanStr = Mid(iScanStr, i + 1)
'                End If
'            End Select
'        End If
'    Loop
'    ret = iSecurPharmValues(0)
'
'
'
''    If (ret <> "") Then
'    If (iSecurPharmValues(0) <> "") Then
'        Dim lRecs&
'
'        iSecurPharmDMC = ScanStr
'
'        If (Mid(iSecurPharmValues(3), 5, 2) = "00") Then
'            Mid(iSecurPharmValues(3), 5, 2) = "01"
'        End If
'
'        h = iSecurPharmValues(3)
''        h = Strings.Left(h, 2) + "." + Mid(h, 3, 2) + "." + "20" + Mid(h, 5, 2)
'        h = Mid(h, 5, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 1, 2)
'
'        SQLStr = "SELECT * FROM QR_SecurPharm WHERE "
'        SQLStr = SQLStr + "(Pzn=" + ret + ") AND (SerienNr='" + iSecurPharmValues(2) + "') AND (Charge='" + iSecurPharmValues(1) + "') AND (Verfall='" + h + "')"
'        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr, 0)
'        If (FabsErrf = 0) Then
'            If (ArtikelRec!AbholNr = 0) And (AbholNr > 0) Then
'                SQLStr = "UPDATE QR_SecurPharm SET AbholNr=" + CStr(AbholNr) + " WHERE "
'                SQLStr = SQLStr + "(Pzn=" + ret + ") AND (SerienNr='" + iSecurPharmValues(2) + "') AND (Charge='" + iSecurPharmValues(1) + "') AND (Verfall='" + h + "')"
'                Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
'            End If
'
'            Dim iLiefDat As Integer
'            iLiefDat = 0
'            Dim dtLieferDatum As Date
'            dtLieferDatum = CheckNullDate(ArtikelRec!LieferDatum)
'            If (Year(dtLieferDatum) > 2000) Then
'                hh = Format(dtLieferDatum, "DDMMYY")
'                iLiefDat = iDate(hh)
'            End If
'            iSecurPharmLieferDatum = iLiefDat
'
'            iSecurPharmLiefNr = CheckNullLong(ArtikelRec!LiefNr)
'            iSecurPharmBeleg = CheckNullStr(ArtikelRec!Beleg)
'            iSecurPharmAbholNr = CheckNullLong(ArtikelRec!AbholNr)
'        Else
'            SQLStr = "INSERT INTO QR_SecurPharm (Pzn,SerienNr,Verfall,Charge,LieferDatum,LiefNr,Beleg,AbholNr,ProduktCode,DMC"
'            SQLStr = SQLStr + ") VALUES ("
'    '        SQLStr = SQLStr + ret + ",'" + iSecurPharmValues(2) + "','" + h + "','" + iSecurPharmValues(1) + "','" + Format(Now, "dd.MM.yyyy hh:mm:ss") + "'," + "8" + ",'" + "1234567" + "'"
'            SQLStr = SQLStr + ret + ",'" + iSecurPharmValues(2) + "','" + h + "','" + iSecurPharmValues(1) + "','" + LieferDatum + "'," + CStr(ActBelegLief) + ",'" + ActBeleg + "'," + CStr(AbholNr) + ",'" + "','" + ScanStr + "'"
'            SQLStr = SQLStr + ")"
'            Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
'
'            SQLStr = "UPDATE Artikel SET QR_SecurPharm=1"
'            SQLStr = SQLStr + " WHERE Pzn=" + ret
'            Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
'
'            iSecurPharmLieferDatum = ActBelegDatum
'            iSecurPharmLiefNr = ActBelegLief
'            iSecurPharmBeleg = ActBeleg
'            iSecurPharmAbholNr = AbholNr
'
'            iSecurPharmNeu = True
'        End If
'    End If
'
'End If
'
'CheckSecurPharm = ret
'
'Call clsError.DefErrPop
'End Function

Function CheckSecurPharm$(ScanStr$, Optional ActBeleg$ = "", Optional ActBelegDatum% = 0, Optional ActBelegLief% = 0, Optional AbholNr% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckSecurPharm$")
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
Dim i, ind, l, iOk, gef As Integer
Dim h$, hh$, iScanStr$, ret$

Dim MessageHeader$, MessageTrailer$, FormatHeader$, FormatTrailer$, FieldSeperator$, RecordSeperator$, sField$, sValue$, SQLStr$
Dim ASC_06Macro_Header$, ASC_06Macro_Trailer$
Dim ASC_DataIdentifier$()

Dim GS1_Header$
Dim GS1_DataIdentifier$()

Dim LieferDatum As String
LieferDatum = ""
If (ActBelegDatum > 0) Then
    h = sDate(ActBelegDatum)
    LieferDatum = Mid(h, 1, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5, 2)
End If

'IFA_DataIdentifier = Array("9N", "1T", "S", "D", "16D", "8P")
ASC_DataIdentifier = Split("9N|1T|S|D|16D|8P", "|")
GS1_DataIdentifier = Split("01|10|21|17|16D|01", "|")

iSecurPharmFields = Split("Pzn|Charge|SerienNr|Verfalldatum|Herstelldatum|GTIN", "|")
iSecurPharmValues = Split("|||||", "|")
iSecurPharmNeu = 0
iSecurPharmDMC = ""
iSecurPharmLieferDatum = 0
iSecurPharmLiefNr = 0
iSecurPharmBeleg = ""
iSecurPharmAbholNr = 0


FieldSeperator = Chr(29)    'GS
RecordSeperator = Chr(30)   'RS
MessageTrailer = Chr(4) 'EOT

FieldSeperator = "|"    'GS
RecordSeperator = "{"  'RS
MessageTrailer = "}" 'EOT

h$ = Space(5)
l = GetPrivateProfileString("Codes", "GS", "|", h$, 6, CurDir + "\SecurPharm.ini")
FieldSeperator = Left$(h$, l)
    
h$ = Space(5)
l = GetPrivateProfileString("Codes", "RS", "{", h$, 6, CurDir + "\SecurPharm.ini")
RecordSeperator = Left$(h$, l)
    
h$ = Space(5)
l = GetPrivateProfileString("Codes", "EOT", "}", h$, 6, CurDir + "\SecurPharm.ini")
MessageTrailer = Left$(h$, l)
    



MessageHeader = "[)>" + RecordSeperator
FormatHeader = "06" + FieldSeperator
FormatTrailer = RecordSeperator

ASC_06Macro_Header = MessageHeader + FormatHeader
ASC_06Macro_Trailer = FormatTrailer + MessageTrailer

GS1_Header = GS1_DataIdentifier(0)

'For i = 0 To UBound(SecurPharmFields)
'    SecurPharmFields(i) = ""
'Next i

ScanStr$ = Trim(ScanStr)

iScanStr = ScanStr
ret = iScanStr

l = Len(iScanStr)

If (l > 20) Then
    If (ScanStr = "0104150123456789012345678901") Then
        MsgBox (ScanStr + vbCrLf + "PZN: 07765007")
        iSecurPharmValues(0) = "07765007"
        iSecurPharmValues(1) = "CH12345678"
        iSecurPharmValues(2) = "S07765007"
        iSecurPharmValues(3) = "241231"
    Else
        If Right(iScanStr, 1) = "{" Or Right(iScanStr, 1) = "}" Then iScanStr = Left(iScanStr, Len(iScanStr) - 1)
        If Right(iScanStr, 1) = "{" Or Right(iScanStr, 1) = "}" Then iScanStr = Left(iScanStr, Len(iScanStr) - 1)
        i = InStr(iScanStr, "{")
        If i > 0 Then iScanStr = Mid(iScanStr, i + 1)
    
        Dim IstASC As Boolean
        IstASC = False
    
        i = InStr(UCase(Left(iScanStr, 10)), "9N")
        If i = 0 Then   'GS1
            i = InStr(iScanStr, "0104150")
            If i = 0 Then
                If Left(iScanStr, 2) = "01" Then
                    'ausländischer GTIN
                    'iScanStr = Mid(iScanStr, 17)
                    i = 1
                Else
                    i = InStr(iScanStr, "9N11")
                    If i = 0 Then
                        CheckSecurPharm = ret
                        Call clsError.DefErrPop: Exit Function
                    Else
                        IstASC = True
                        iSecurPharmValues(0) = Mid(iScanStr, i + 4, 8)
                        iScanStr = Left(iScanStr, i - 1) + Mid(iScanStr, i + 15)
                    End If
                End If
            End If
            If Not IstASC Then
                If i > 0 And i < 7 Then
                    iScanStr = Mid(iScanStr, i + 7)   '014150 weg
                    iSecurPharmValues(0) = Left(iScanStr, 8)
                    iScanStr = Mid(iScanStr, 10)    '1 PZ
                ElseIf i > 0 Then
                    iSecurPharmValues(0) = Mid(iScanStr, i + 7, 8)
                    iScanStr = Left(iScanStr, i - 1) + Mid(iScanStr, i + 16)
                End If
            End If
        Else    'ASC
            IstASC = True
    '        If (Right(iScanStr, Len(ASC_06Macro_Trailer)) = ASC_06Macro_Trailer) Then
    '            iScanStr = Left(iScanStr, Len(iScanStr) - (Len(ASC_06Macro_Trailer)))
    '        End If
            iScanStr = Mid(iScanStr, i + 4)   '9N11 weg
            iSecurPharmValues(0) = Left(iScanStr, 8)
            iScanStr = Mid(iScanStr, 11)    '2PZ
        End If
    
    '    ret = 1
        Do While Len(iScanStr) > 1
            iScanStr = Trim(iScanStr)
            If Left(iScanStr, 1) = "|" Or Left(iScanStr, 1) = "<" Then iScanStr = Trim(Mid(iScanStr, 2))
            If IstASC Then
                If UCase(Left(iScanStr, 2)) = "1T" Then    'Charge
                    iScanStr = Mid(iScanStr, 3)
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = InStr(iScanStr, "<")
                    If i = 0 Then i = Len(iScanStr) + 1
                    iSecurPharmValues(1) = Trim(Left(iScanStr, i - 1))
                    iScanStr = Mid(iScanStr, i + 1)
                ElseIf UCase(Left(iScanStr, 1)) = "S" Then   'Seriennummer
                    iScanStr = Mid(iScanStr, 2)
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = InStr(iScanStr, "<")
                    If i = 0 Then i = Len(iScanStr) + 1
                    iSecurPharmValues(2) = Trim(Left(iScanStr, i - 1))
                    iScanStr = Mid(iScanStr, i + 1)
                ElseIf UCase(Left(iScanStr, 1)) = "D" Then   'Verfall
                    iSecurPharmValues(3) = Mid(iScanStr, 2, 6)
                    iScanStr = Mid(iScanStr, 8)
                Else
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = InStr(iScanStr, "<")
                    If i = 0 Then i = Len(iScanStr) + 1
                    If i > 0 Then
                        iScanStr = Mid(iScanStr, i + 1)
                    End If
                End If
            Else
                Select Case Left(iScanStr, 2)
                Case "21"   'Seriennummer
                    iScanStr = Mid(iScanStr, 3)
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = Len(iScanStr) + 1
                    iSecurPharmValues(2) = Left(iScanStr, i - 1)
                    iScanStr = Mid(iScanStr, i + 1)
                Case "10"   'Charge
                    iScanStr = Mid(iScanStr, 3)
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = Len(iScanStr) + 1
                    iSecurPharmValues(1) = Left(iScanStr, i - 1)
                    iScanStr = Mid(iScanStr, i + 1)
                Case "17"   'Verfall
                    iSecurPharmValues(3) = Mid(iScanStr, 3, 6)
                    iScanStr = Mid(iScanStr, 9)
                Case "71"   'PZN bei ausländischem GTIN
    '                If Val(iSecurPharmValues(0)) = 0 And Left(iScanStr, 3) = "710" Then
                    If Left(iScanStr, 3) = "710" Then
                        iSecurPharmValues(0) = Mid(iScanStr, 4, 8)
                        iScanStr = Mid(iScanStr, 12)
                        If Left(iScanStr, 1) = "|" Then iScanStr = Mid(iScanStr, 2)
                    Else
                        i = InStr(iScanStr, "|")
                        If i = 0 Then i = Len(iScanStr) + 1
                        If i > 0 Then iScanStr = Mid(iScanStr, i + 1)
                    End If
                Case Else
                    i = InStr(iScanStr, "|")
                    If i = 0 Then i = Len(iScanStr) + 1
                    If i > 0 Then
                        iScanStr = Mid(iScanStr, i + 1)
                    End If
                End Select
            End If
        Loop
    End If
    
    ret = iSecurPharmValues(0)


'    If (ret <> "") Then
    If (iSecurPharmValues(0) <> "") Then
        Dim lRecs&
        
        iSecurPharmDMC = ScanStr
        
        If (Mid(iSecurPharmValues(3), 5, 2) = "00") Then
            Mid(iSecurPharmValues(3), 5, 2) = "01"
        End If

        h = iSecurPharmValues(3)
'        h = Strings.Left(h, 2) + "." + Mid(h, 3, 2) + "." + "20" + Mid(h, 5, 2)
        h = Mid(h, 5, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 1, 2)
        

        SQLStr = "SELECT * FROM QR_SecurPharm WHERE "
        SQLStr = SQLStr + "(Pzn=" + ret + ") AND (SerienNr='" + iSecurPharmValues(2) + "') AND (Charge='" + iSecurPharmValues(1) + "') AND (Verfall='" + h + "')"
        FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr, 0)
        If (FabsErrf = 0) Then
            If (ArtikelRec!AbholNr = 0) And (AbholNr > 0) Then
                SQLStr = "UPDATE QR_SecurPharm SET AbholNr=" + CStr(AbholNr) + " WHERE "
                SQLStr = SQLStr + "(Pzn=" + ret + ") AND (SerienNr='" + iSecurPharmValues(2) + "') AND (Charge='" + iSecurPharmValues(1) + "') AND (Verfall='" + h + "')"
                Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
            End If
            
            Dim iLiefDat As Integer
            iLiefDat = 0
            Dim dtLieferDatum As Date
            dtLieferDatum = CheckNullDate(ArtikelRec!LieferDatum)
            If (Year(dtLieferDatum) > 2000) Then
                hh = Format(dtLieferDatum, "DDMMYY")
                iLiefDat = iDate(hh)
            End If
            iSecurPharmLieferDatum = iLiefDat
            
            iSecurPharmLiefNr = CheckNullLong(ArtikelRec!LiefNr)
            iSecurPharmBeleg = CheckNullStr(ArtikelRec!Beleg)
            iSecurPharmAbholNr = CheckNullLong(ArtikelRec!AbholNr)
        Else
            SQLStr = "INSERT INTO QR_SecurPharm (Pzn,SerienNr,Verfall,Charge,LieferDatum,LiefNr,Beleg,AbholNr,ProduktCode,DMC"
            SQLStr = SQLStr + ") VALUES ("
    '        SQLStr = SQLStr + ret + ",'" + iSecurPharmValues(2) + "','" + h + "','" + iSecurPharmValues(1) + "','" + Format(Now, "dd.MM.yyyy hh:mm:ss") + "'," + "8" + ",'" + "1234567" + "'"
            SQLStr = SQLStr + ret + ",'" + iSecurPharmValues(2) + "','" + h + "','" + iSecurPharmValues(1) + "','" + LieferDatum + "'," + CStr(ActBelegLief) + ",'" + ActBeleg + "'," + CStr(AbholNr) + ",'" + "','" + ScanStr + "'"
            SQLStr = SQLStr + ")"
            Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
        
            SQLStr = "UPDATE Artikel SET QR_SecurPharm=1"
            SQLStr = SQLStr + " WHERE Pzn=" + ret
            Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs&, adExecuteNoRecords)
            
            iSecurPharmLieferDatum = ActBelegDatum
            iSecurPharmLiefNr = ActBelegLief
            iSecurPharmBeleg = ActBeleg
            iSecurPharmAbholNr = AbholNr
            
            iSecurPharmNeu = True
        End If
    End If

End If

CheckSecurPharm = ret

Call clsError.DefErrPop
End Function

Function GetTaskId(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetTaskId&")
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
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim x As Long
Dim ThreadID As Long

TaskName = UCase(TaskName)

GetTaskId = 0
'CurrWnd = GetWindow(GetDesktopWindow, GW_HWNDFIRST)
CurrWnd = GetWindow(GetDesktopWindow, GW_CHILD)
Do While CurrWnd <> 0
  Length = GetWindowTextLength(CurrWnd)
  ListItem = Space(Length + 1)
  Length = GetWindowText(CurrWnd, ListItem, Length + 1)
  If Length > 0 Then
    x = GetWindowThreadProcessId(CurrWnd, ThreadID)
    If UCase(Left(ListItem, Len(TaskName))) = TaskName Then
      GetTaskId = ThreadID
      Exit Do
    End If
  End If
  CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call clsError.DefErrPop
End Function

Function FileExist%(sDatei$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FileExist%")
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
Dim SearchHandle&
Dim FindDataRec As WIN32_FIND_DATA

SearchHandle& = FindFirstFile(sDatei$, FindDataRec)
FileExist% = (SearchHandle& <> INVALID_HANDLE_VALUE)

Call clsError.DefErrPop
End Function

Function IsLetter(ByVal character As String) As Boolean
    IsLetter = UCase$(character) <> LCase$(character)
End Function



