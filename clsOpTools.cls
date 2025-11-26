VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpTools"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function CVI%(X$)
Dim temp%

CopyMemory temp%, ByVal X$, 2
CVI% = temp%

End Function

Function CVL&(X$)
Dim temp&

CopyMemory temp&, ByVal X$, 4
CVL& = temp&

End Function

Function CVS!(X$)
Dim temp!

CopyMemory temp!, ByVal X$, 4
CVS! = temp!

End Function

Function CVD#(X$)
Dim temp#

CopyMemory temp#, ByVal X$, 8
Call DxToIEEEd(temp#)
CVD# = temp#

End Function

Function CVDat%(cDatum$)
Dim d$

d$ = Right$(cDatum$, 1) + Left$(cDatum$, 1)
CVDat% = CVI%(d$)

End Function

Function CVDatum$(cDatum$)
Dim d$

d$ = Right$(cDatum$, 1) + Left$(cDatum$, 1)
d$ = sDate$(CVI%(d$))
If Len(d$) < 6 Then
    d$ = Space$(8)
Else
    d$ = DateLong$(d$)
End If
CVDatum$ = d$

End Function

Function CVDatum2$(ByVal iDatum%)
Dim d$

d$ = sDate$(iDatum%)
If Len(d$) < 6 Then
    d$ = Space$(8)
Else
    d$ = DateLong$(d$)
End If
CVDatum2$ = d$

End Function

Function sDate$(iDatum%)
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

End Function

Function DateLong$(sDatum$)
Dim yy%

yy% = Val(Right$(sDatum$, 2))
If yy% >= 50 Then yy% = yy% + 1900 Else yy% = yy% + 2000
DateLong$ = Right$(Str$(yy%), 4) + Mid$(sDatum$, 3, 2) + Left$(sDatum$, 2)

End Function

Function MKI$(X%)
Dim temp$

temp$ = String(2, 0)
CopyMemory ByVal temp$, X%, 2
MKI$ = temp$

End Function

Function MKL$(X&)
Dim temp$

temp$ = String(4, 0)
CopyMemory ByVal temp$, X&, 4
MKL$ = temp$

End Function

Function MKS$(X!)
Dim temp$

temp$ = String(4, 0)
CopyMemory ByVal temp$, X!, 4
MKS$ = temp$

End Function

Function MKD$(X#)
Dim temp$

temp$ = String(8, 0)
CopyMemory ByVal temp$, X#, 4
MKD$ = temp$

End Function

Function MKDate$(iDatum%)
Dim d$

d$ = MKI$(iDatum%)
MKDate$ = Right$(d$, 1) + Left$(d$, 1)

End Function

Function iDate%(sDatum$)
Dim X%, xd%, xm%, xy%
Dim xdatum$

'* 1.1.1972 (= 1) bis 31.12.2049 (=28490)
X% = 0: xdatum$ = sDatum$
If (Len(xdatum$) = 6) Then
    Do
        X% = InStr(xdatum$, " ")
        If (X% > 0) Then Mid$(xdatum$, X%, 1) = "0"
    Loop While X% > 0
    xd% = Val(Mid$(xdatum$, 1, 2))
    xm% = Val(Mid$(xdatum$, 3, 2))
    xy% = Val(Mid$(xdatum$, 5, 2))
    If (xy% < 50) Then xy% = xy% + 2000 Else xy% = xy% + 1900
    X% = Int(365.25 * (xy% - 1972) + 0.75) + Int(30.55 * xm% - 29.95 + 0.001)
    X% = X% + 2 * (xm% > 2) + xd% - ((xm% > 2) And xy% / 4 = Int(xy% / 4))
    If (xdatum$ <> sDate$(X%)) Then X% = 0
End If
iDate% = X%
End Function

Function HoleKundenName$(KundenNr%)
Dim clsKunden1 As New clsKunden
Dim h$

clsKunden1.OpenDatei
clsKunden1.GetRecord (KundenNr% + 1)
h$ = RTrim$(clsKunden1.wert("name", 0))
Set clsKunden1 = Nothing

HoleKundenName$ = h$

End Function

Function Bcd2ascii(ByVal bcd As String, stellen As Integer) As String
Dim i As Integer
Dim ascii As String
'ascii = Ergebnis Zahlenstring
'bcd = binärcodierte Zahl
'stellen = Anzahl der Stellen in Ergebnisfeld
'z.B.: stellen = 3   bcd = chr(&H45) + chr(&H67) => ascii = "456"

ascii = String(stellen, "0")
If bcd <> String(Len(bcd), "*") Then
  For i = 1 To stellen Step 2
    Mid(ascii, i, 1) = Chr(Int(Asc(Mid(bcd, Int(i / 2) + 1, 1)) / 16) + &H30)
    If i < stellen Then
      Mid(ascii, i + 1, 1) = Chr((Asc(Mid(bcd, Int(i / 2) + 1, 1)) Mod 16) + &H30)
    End If
  Next i
End If
Bcd2ascii = ascii

End Function

Function Ascii2bcd(ByVal ascii As String, stellen As Integer) As String
Dim i As Integer
Dim bcd As String
Dim bcdwert As Integer

ascii = Right$(String$(stellen, "0") + ascii, stellen)
For i = 1 To stellen
  If InStr("0123456789", Mid$(ascii, i, 1)) = 0 Then
    Mid$(ascii, i, 1) = "0"
  End If
Next i

bcd = String$(Int((stellen + 1)) / 2, 0)
For i = 1 To stellen Step 2
  bcdwert = (Asc(Mid(ascii, i, 1)) - &H30) * 16
  If i < stellen Then
    bcdwert = bcdwert + Asc(Mid$(ascii, i + 1, 1)) - &H30
  End If
  Mid$(bcd$, Int(i / 2) + 1, 1) = Chr$(bcdwert)
Next i

Ascii2bcd = bcd

End Function

Sub EanPruef(EAN As String)
Dim i As Integer, Pruefsumme As Integer, PruefZiffer As Integer

If Val(EAN) > 0 And Len(EAN) = 13 Then
  Pruefsumme = 0
  For i = 1 To 12
    Pruefsumme = Pruefsumme + Val(Mid(EAN, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
  Next i
  PruefZiffer = 10 - (Pruefsumme Mod 10)
  Mid(EAN, 13, 1) = Right(Str(PruefZiffer), 1)
End If

End Sub

Sub Taxe2ast(pzn$)
Dim i%, wg%, Appli%, nAm%
Dim dPreis#
Dim SQLStr$, h$, h2$, X$, ascii$, gc$, mw$, Rez$, arez$
Dim recTaxe As Recordset

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
If (recTaxe.EOF = True) Then Exit Sub


'Call WGconvEinlesen

'* pzn kurz meng
Ast1.pzn = Format$(recTaxe!pzn, "0000000")
Ast1.kurz = Left$(Left$(recTaxe!Name, 28) + Space$(28), 28)
Ast1.meng = Right$(recTaxe!menge, 5)

'* meh
Ast1.meh = recTaxe!Einheit

'* aep
dPreis# = recTaxe!EK / 100
'Call DxToMBFd(dPreis#)
Ast1.AEP = dPreis#

'* Festbetrag
dPreis# = recTaxe!FESTBETRAG / 100
'Call DxToMBFd(dPreis#)
Ast1.kp = dPreis#

'* avp
dPreis# = recTaxe!VK / 100
'Call DxToMBFd(dPreis#)
Ast1.AVP = dPreis#

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
If (wg% >= 10) And (wg% <= 19) Then
  Appli% = -1
  wg% = wg% - 10
End If
Ast1.wg = WGconvert(wg%)

'* herst
Ast1.herst = recTaxe!HerstellerKB

'* lic
h$ = recTaxe!UeberGH
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
h$ = Chr$(recTaxe!Lagerung)
If (Val(h$) = 0) Then h$ = " "
Ast1.lac = h$


'* rez
Rez$ = Chr$(recTaxe!AbgabeBest)
nAm% = 0: If (InStr("23", Rez$) <> 0) Then nAm% = -1
If (Appli%) Then
    arez$ = "A ": If (nAm%) Then arez$ = "AP"
ElseIf (recTaxe!BTMKz) Then
    arez$ = "SG"
ElseIf (InStr("156", Rez$) <> 0) Then
    arez$ = "+ "
ElseIf nAm% Then
    'Nicht-Arzneimittel oder nicht apothekenpflichtiges Arzneimittel
    arez$ = "P "
Else
    arez$ = "  "
End If
Ast1.Rez = arez$

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
Ast1.vnr = Space$(8)

'* kz
Ast1.kz = " "

'* abl
h$ = " "
If (recTaxe!VerfallArt = "1") Then
  h$ = "A"
End If
Ast1.abl = h$

Ast1.frei = Chr$(0)

End Sub

Function WGconvert$(wg%)
Dim NeuWg%, iw%
Dim TestWg$, X$

NeuWg% = 0: TestWg$ = "," + Mid$(Str$(wg%), 2) + ","
For iw% = 1 To 9
  If (NeuWg% = 0) Then
    X$ = Para1.WGconv(iw%): If (InStr(X$, TestWg$) <> 0) Then NeuWg% = iw%
  End If
Next iw%
If (NeuWg% = 0) Then NeuWg% = 9

WGconvert$ = Right$(Str$(NeuWg%), 1) + " "

End Function

Sub WarteAufTaskEnde(SuchTaskId&, pForm As Object)
Dim erg%

pForm.Enabled = False
Do
    erg% = CheckTask%(SuchTaskId&, pForm)
    If (erg% = False) Then Exit Do
    DoEvents
Loop
pForm.Enabled = True

End Sub

Function CheckTask%(SuchTaskId&, pForm As Object)
Dim ThreadID
Dim CurrWnd&, X&

CheckTask% = True

CurrWnd& = GetWindow(pForm.hWnd, GW_HWNDFIRST)
While (CurrWnd& <> 0)
  X& = GetWindowThreadProcessId(CurrWnd&, ThreadID)
  If (ThreadID = SuchTaskId&) Then Exit Function
  CurrWnd& = GetWindow(CurrWnd&, GW_HWNDNEXT)
Wend

CheckTask% = False
End Function

