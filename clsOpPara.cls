VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iZuzWert!(6)

Dim iFistam$(1)
Dim iBenutz$

Dim iPersonal$(50)

Dim iTaxeLw$
Dim iCdLw$
Dim iLwK$

Dim iBestellPeriode%
Dim iLagerkosten#
Dim iPersonalKosten#
Dim iTfaktor#
Dim iFakNatu#
Dim iFakBar#

Dim iPD1$, iPD2$, iPD3$
Dim iPP1$, iPP2$, iPP3$
Dim iaText1$, iaText2$, iaText3$
Dim iDAFUE$
Dim iDP03File$
Dim iPort$
Dim ipDrucker1%, ipDrucker2%, ipDrucker3%
Dim ipPreis3%

Dim iMitKz%
Dim ipPreis%
Dim ipplr%
Dim ipplz%
Dim iFixEti%
Dim iText2%
Dim iPpreis2%
Dim iLeer%
Dim ijp$
Dim ifText1$
Dim ifText2$
Dim ifText3$

Dim iTaraUser$
Dim iFiliUser$
Dim iRezRech$
Dim iPosAktiv$
Dim iDABM%
Dim ibp1%
Dim iBP2%
Dim iSPREIZ%
Dim iSTAFF2%
Dim iLager1%
Dim iLager2%
Dim iTestZwei$
Dim ixBP$
Dim iiBP!
Dim iAblK$
Dim iTOT!

Dim ibbm%(10)
Dim imm%(10)
Dim iBMS!(10)
Dim iPosAktivWG$
Dim iDirektLieferanten$
Dim iLangsamLM%
Dim iLagerAuffuellen%
Dim iTaraKontrolle%
Dim iMonNBest%
Dim iLstGemeinsam$
Dim iBVRetour$

Dim iisdn%

Dim iWGconv$(9)


Dim PdateiFix As String * 30
Dim PDATEI%

Private Const DefErrModul = "clsOpPara.cls"

Private Sub Class_Initialize()
Set Para1 = Me
End Sub

Sub HoleZuzahlungen()
Dim i%, PARAM%
Dim s$

PARAM% = clsDat.FileOpen("param.dat", "I")
If (PARAM% > 0) Then
    For i% = 0 To 18
        Line Input #PARAM%, s$
    Next i%
    
    Line Input #PARAM%, s$: iZuzWert!(4) = Val(s$)
    Line Input #PARAM%, s$: iZuzWert!(3) = Val(s$)
    Line Input #PARAM%, s$: iZuzWert!(5) = Val(s$)
    Line Input #PARAM%, s$: iZuzWert!(6) = Val(s$)
    
    Close #PARAM%
End If

End Sub

Sub HoleFirmenStamm()
Dim FISTAM%, erg%
Dim s$, buf$

s$ = "fistam.dat"
FISTAM% = clsDat.FileOpen(s$, "R")

buf$ = Space$(45)
Get FISTAM%, , buf$
Get FISTAM%, , buf$
iFistam$(0) = RTrim$(buf$)
Call OemToChar(iFistam$(0), iFistam$(0))
Get FISTAM%, , buf$
iFistam$(1) = RTrim$(buf$)
Call OemToChar(iFistam$(1), iFistam$(1))
Get FISTAM%, , buf$
iBenutz$ = RTrim$(buf$)

Close FISTAM%

End Sub

Sub EinlesenPersonal()
Dim pass As New clsPasswort
Dim i%
Dim h$

pass.OpenDatei

For i% = 1 To 50
    pass.GetRecord (i% + 1)
    iPersonal$(i%) = pass.Name
Next i%

pass.CloseDatei
pass.FreeClass

End Sub

Sub AuslesenPdatei()
Dim s$

s$ = "pdatei.dat"
PDATEI% = clsDat.FileOpen(s$, "R", "R", Len(PdateiFix))
If (PDATEI% > 0) Then
    Call DecodPfade
    Call LagerParameter
    Call ReadPara
    Call WGconvEinlesen
    
    iisdn% = False
    Get #PDATEI%, 46, PdateiFix
    If (Mid$(PdateiFix, 8, 1) = "J") Then iisdn% = True
    
    Close #PDATEI%
End If

End Sub

Sub DecodPfade()

iTaxeLw$ = "Y"
iCdLw$ = "Y"
iLwK$ = "K"

Get #PDATEI%, 30, PdateiFix
iTaxeLw$ = Mid$(PdateiFix, 25, 1)

Get #PDATEI%, 43, PdateiFix
iCdLw$ = Mid$(PdateiFix, 11, 1)
iLwK$ = Mid$(PdateiFix, 13, 1)

End Sub

Sub LagerParameter()
Dim h$

Get #PDATEI%, 2, PdateiFix

h$ = PdateiFix

iBestellPeriode% = clsOpTool.CVI%(Mid$(h$, 19))
iLagerkosten# = Asc(Mid$(h$, 25, 1))
iPersonalKosten# = Asc(Mid$(h$, 26, 1))
iTfaktor# = Asc(Mid$(h$, 27, 1))
iFakNatu# = CSng(Asc(Mid$(h$, 28, 1))) / 100
iFakBar# = CSng(Asc(Mid$(h$, 29, 1))) / 100

End Sub

Sub ReadPara()
Dim j%, erg%
Dim xs!
Dim l&
Dim s$, h$, X$

Get #PDATEI%, 1, PdateiFix: iPD1$ = PdateiFix
Get #PDATEI%, 2, PdateiFix$: iPD2$ = PdateiFix
Get #PDATEI%, 3, PdateiFix$: iPD3$ = PdateiFix

Get #PDATEI%, 16, PdateiFix
iPP1$ = Left$(PdateiFix, 10): iPP2$ = Mid$(PdateiFix, 11, 10): iPP3$ = Mid$(PdateiFix, 21, 10)
Get #PDATEI%, 17, PdateiFix
iaText1$ = PdateiFix
Get #PDATEI%, 18, PdateiFix
iaText2$ = PdateiFix
Get #PDATEI%, 19, PdateiFix
iaText3$ = PdateiFix
Get #PDATEI%, 20, PdateiFix
iDAFUE$ = Left$(PdateiFix, 3)
Get #PDATEI%, 21, PdateiFix
X$ = Left$(PdateiFix, 30): X$ = LTrim$(X$): X$ = RTrim$(X$): iDP03File$ = X$
Get #PDATEI%, 22, PdateiFix
iPort$ = Left$(PdateiFix, 3)

Get #PDATEI%, 23, PdateiFix
ipDrucker1% = Asc(Mid$(PdateiFix, 1, 1))
If ((ipDrucker1% < 1) Or (ipDrucker1% > 9)) Then ipDrucker1% = 1
ipDrucker2% = Asc(Mid$(PdateiFix, 2, 1))
If ((ipDrucker2% < 1) Or (ipDrucker2% > 9)) Then ipDrucker2% = 1
ipDrucker3% = Asc(Mid$(PdateiFix, 3, 1))
If ((ipDrucker3% < 1) Or (ipDrucker3% > 9)) Then ipDrucker3% = 1
ipPreis3% = Asc(Mid$(PdateiFix, 4, 1))
If ((ipPreis3% < 0) Or (ipPreis3% > 4)) Then ipPreis3% = 0
If (ipPreis3% = 4) Then ipPreis3% = 3: iMitKz% = True

Get #PDATEI%, 24, PdateiFix
ipPreis% = Val(Left$(PdateiFix, 1))
ipplr% = Asc(Mid$(PdateiFix, 2, 1))
If (ipplr% = 0) Then ipplr% = 3
ipplz% = Asc(Mid$(PdateiFix, 3, 1))
iFixEti% = Val(Mid$(PdateiFix, 4, 1))
iText2% = Val(Mid$(PdateiFix, 6, 1)): iPpreis2% = Val(Mid$(PdateiFix, 7, 1))
iLeer% = Val(Mid$(PdateiFix, 8, 1))

Get #PDATEI%, 25, PdateiFix
ijp$ = PdateiFix
Get #PDATEI%, 26, PdateiFix
ifText1$ = PdateiFix
Get #PDATEI%, 27, PdateiFix
ifText2$ = PdateiFix
Get #PDATEI%, 28, PdateiFix
ifText3$ = PdateiFix


Get #PDATEI%, 30, PdateiFix
iTaraUser$ = Mid$(PdateiFix, 1, 10): iFiliUser$ = Mid$(PdateiFix, 11, 10)
iRezRech$ = Mid$(PdateiFix, 21, 1): iPosAktiv$ = Mid$(PdateiFix, 24, 1)
iDABM% = Asc(Right$(iPD2$, 1))
ibp1% = clsOpTool.CVI%(Mid$(iPD2$, 19)): iBP2 = clsOpTool.CVI%(Mid$(iPD2$, 21))
iSPREIZ = Asc(Mid$(iPD2$, 23)): iSTAFF2 = Asc(Mid$(iPD2$, 24))
iLager1 = Asc(Mid$(iPD2$, 25)): iLager2 = Asc(Mid$(iPD2$, 26))
iTestZwei$ = Mid$(iPD3$, 17, 1): If InStr("TS", iTestZwei$) = 0 Then iTestZwei$ = "T"
ixBP$ = Mid$(iPD3$, 20, 1): If InStr("123456789", ixBP$) = 0 Then ixBP$ = "4"
iiBP! = Val(ixBP$): iAblK$ = Mid$(iPD3$, 21, 1): iTOT = iSPREIZ / 100
For j% = 1 To 10
    ibbm%(j%) = clsOpTool.CVI%(Mid$(iPD1$, (j% - 1) * 3 + 1, 2))
    imm%(j%) = Asc(Mid$(iPD1$, (j% - 1) * 3 + 3, 1))
    xs! = Abs(((ibbm%(j%) - ibbm%(j% - 1)) = 0) * iTOT * ibbm%(j%))
    If (xs! = 0!) Then
        iBMS!(j%) = ibbm%(j%)
    Else
        iBMS!(j% - 1) = ibbm%(j% - 1) - xs! / 2
        iBMS!(j%) = ibbm%(j%) + xs! / 2
    End If
Next j%

Get #PDATEI%, 43, PdateiFix
iPosAktivWG$ = Mid$(PdateiFix, 1, 10)
Get #PDATEI%, 45, PdateiFix
iDirektLieferanten$ = Mid$(PdateiFix, 2, 24)

Get #PDATEI%, 46, PdateiFix
iLangsamLM% = Asc(Mid$(PdateiFix, 9, 1))
iLagerAuffuellen% = 0: iTaraKontrolle% = 0:
If (InStr("Jj", Mid$(PdateiFix, 10, 1)) <> 0) Then iLagerAuffuellen% = True
If (InStr("Jj", Mid$(PdateiFix, 12, 1)) <> 0) Then iTaraKontrolle% = True
iMonNBest% = Val(Mid$(PdateiFix, 18, 2))    '1.58
If ((iMonNBest% < 1) Or (iMonNBest% > 24)) Then iMonNBest% = 6    '1.58
iMonNBest% = iMonNBest% * 30.33     '1.58
iLstGemeinsam$ = "J": If Mid$(PdateiFix, 23, 1) = "N" Then iLstGemeinsam$ = "N" '1.61

'If land$ = "D" Then
    iLagerAuffuellen% = True
    iBVRetour$ = Space$(10)
    l& = GetPrivateProfileString("Warenwirtschaft", "BestVorsRetourware", " ", iBVRetour$, 11, "\user\DP.INI")
    iBVRetour$ = Left$(iBVRetour$, l&)
'   Call IniPar("DP.INI", "WARENWIRTSCHAFT", "BestVorsRetourware", BVRetour$) '1
    If (iBVRetour$ <> "J") Then iBVRetour$ = "N"
'End If
If (InStr(Para1.Benutz, "Q") = 0) Then iTaraKontrolle% = 0
If (iDirektLieferanten$ = String$(Len(iDirektLieferanten$), 0)) Then iDirektLieferanten$ = ""
X$ = iDirektLieferanten$: X$ = LTrim$(X$): X$ = RTrim$(X$): iDirektLieferanten$ = X$

End Sub

Sub WGconvEinlesen()
Dim i%
Dim h$, X$

'Warengruppenkonvertierung einlesen
iWGconv$(0) = "*"

Get #PDATEI%, 30, PdateiFix
'* Warengruppenkonvertierung
For i% = 1 To 9
    Get #PDATEI%, 30 + i%, PdateiFix
    X$ = PdateiFix: X$ = LTrim(X$): X$ = RTrim(X$)
    iWGconv$(i%) = "," + X$ + ","
Next i%

End Sub






Public Property Get ZuzWert!(ind%)
ZuzWert! = iZuzWert!(ind%)
End Property

Public Property Get FISTAM$(ind%)
FISTAM$ = iFistam$(ind%)
End Property

Public Property Get Benutz$()
Benutz$ = iBenutz$
End Property


Public Property Get Personal$(ind%)
Personal$ = iPersonal$(ind%)
End Property


Public Property Get TaxeLw$()
TaxeLw$ = iTaxeLw$
End Property
Public Property Get CdLw$()
CdLw$ = iCdLw$
End Property
Public Property Get LwK$()
LwK$ = iLwK$
End Property

Public Property Get BestellPeriode%()
BestellPeriode% = iBestellPeriode%
End Property
Public Property Get Lagerkosten#()
Lagerkosten# = iLagerkosten#
End Property
Public Property Get PersonalKosten#()
PersonalKosten# = iPersonalKosten#
End Property
Public Property Get Tfaktor#()
Tfaktor# = iTfaktor#
End Property
Public Property Get FakNatu#()
FakNatu# = iFakNatu#
End Property
Public Property Get FakBar#()
FakBar# = iFakBar#
End Property


Public Property Get PD1$()
PD1$ = iPD1$
End Property
Public Property Get PD2$()
PD2$ = iPD2$
End Property
Public Property Get PD3$()
PD3$ = iPD3$
End Property
Public Property Get PP1$()
PP1$ = iPP1$
End Property
Public Property Get PP2$()
PP2$ = iPP2$
End Property
Public Property Get PP3$()
PP3$ = iPP3$
End Property
Public Property Get aText1$()
aText1$ = iaText1$
End Property
Public Property Get aText2$()
aText2$ = iaText2$
End Property
Public Property Get aText3$()
aText3$ = iaText3$
End Property
Public Property Get DAFUE$()
DAFUE$ = iDAFUE$
End Property
Public Property Get DP03File$()
DP03File$ = iDP03File$
End Property
Public Property Get Port$()
Port$ = iPort$
End Property
Public Property Get pDrucker1%()
pDrucker1% = ipDrucker1%
End Property
Public Property Get pDrucker2%()
pDrucker2% = ipDrucker2%
End Property
Public Property Get pDrucker3%()
pDrucker3% = ipDrucker3%
End Property
Public Property Get pPreis3%()
pPreis3% = ipPreis3%
End Property
Public Property Get MitKz%()
MitKz% = iMitKz%
End Property
Public Property Get pPreis%()
pPreis% = ipPreis%
End Property
Public Property Get pplr%()
pplr% = ipplr%
End Property
Public Property Get pplz%()
pplz% = ipplz%
End Property
Public Property Get FixEti%()
FixEti% = iFixEti%
End Property
Public Property Get Text2%()
Text2% = iText2%
End Property
Public Property Get Ppreis2%()
Ppreis2% = iPpreis2%
End Property
Public Property Get Leer%()
Leer% = Leer%
End Property
Public Property Get jp$()
jp$ = ijp$
End Property
Public Property Get fText1$()
fText1$ = ifText1$
End Property
Public Property Get fText2$()
fText2$ = ifText2$
End Property
Public Property Get fText3$()
fText3$ = ifText3$
End Property
Public Property Get TaraUser$()
TaraUser$ = iTaraUser$
End Property
Public Property Get FiliUser$()
FiliUser$ = iFiliUser$
End Property
Public Property Get RezRech$()
RezRech$ = iRezRech$
End Property
Public Property Get PosAktiv$()
PosAktiv$ = iPosAktiv$
End Property
Public Property Get DABM%()
DABM% = iDABM%
End Property
Public Property Get bp1%()
bp1% = ibp1%
End Property
Public Property Get BP2%()
BP2% = iBP2%
End Property
Public Property Get SPREIZ%()
SPREIZ% = iSPREIZ%
End Property
Public Property Get STAFF2%()
STAFF2% = iSTAFF2%
End Property
Public Property Get Lager1%()
Lager1% = iLager1%
End Property
Public Property Get Lager2%()
Lager2% = iLager2%
End Property
Public Property Get TestZwei$()
TestZwei$ = iTestZwei$
End Property
Public Property Get xBP$()
xBP$ = ixBP$
End Property
Public Property Get iBP!()
iBP! = iiBP!
End Property
Public Property Get AblK$()
AblK$ = iAblK$
End Property
Public Property Get TOT!()
TOT! = iTOT!
End Property
Public Property Get bbm%(ind%)
bbm% = ibbm%(ind%)
End Property
Public Property Get mm%(ind%)
mm% = imm%(ind%)
End Property
Public Property Get BMS!(ind%)
BMS! = iBMS!(ind%)
End Property
Public Property Get PosAktivWG$()
PosAktivWG$ = iPosAktivWG$
End Property
Public Property Get DirektLieferanten$()
DirektLieferanten$ = iDirektLieferanten$
End Property
Public Property Get LangsamLM%()
LangsamLM% = iLangsamLM%
End Property
Public Property Get LagerAuffuellen%()
LagerAuffuellen% = iLagerAuffuellen%
End Property
Public Property Get TaraKontrolle%()
TaraKontrolle% = iTaraKontrolle%
End Property
Public Property Get MonNBest%()
MonNBest% = iMonNBest%
End Property
Public Property Get LstGemeinsam$()
LstGemeinsam$ = iLstGemeinsam$
End Property
Public Property Get BVRetour$()
BVRetour$ = iBVRetour$
End Property

Public Property Get isdn%()
isdn% = iisdn%
End Property
Public Property Let isdn(ByVal vNewValue%)
iisdn% = vNewValue%
End Property

Public Property Get WGconv$(ind%)
WGconv$ = iWGconv$(ind%)
End Property

Public Property Get MehrPlatz%()
MehrPlatz% = iMehrPlatz%
End Property

Public Property Get user$()
user$ = iUser$
End Property

