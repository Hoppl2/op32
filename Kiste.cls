VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKiste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const EAN_AUFTRAG = "97"
Const INIFILE_DP = "dp.ini"

Private Type AnfertStruct
    Status As Byte
    VonWann As String * 4
    VonWo As Byte 'String * 1
    Mitarbeiter As String * 1
    Arbeit As String * 5
    Angef As String * 1
    AngefWann As String * 4
    AngefVonWo As String * 1
    ANFBLOCK(0 To 9) As Integer
    KundenNr As Integer
    Profil As Integer
    TextWichtig As Byte
    HatLieferSchein As Byte
    LSnr As Integer
    rest As String * 2
End Type

Private Type AnfWasStruct
    Status As Byte
    WasTun As String * 1
    RezeptNr As Integer
    Preis As String * 9
    ErstMagSatz As Long
    Aconto As String * 9
    RezEan As String * 6
End Type

Private Type AnfInfoStruct
    Protect As Byte 'String * 1
    InfoZeile As String * 24
    RezeptArt As Byte
    rest As String * 6
End Type


Private Type VerkaufStruct
  RezNr       As Integer
  EndeT       As Byte
  pzn         As String * 7
  text        As String * 36
  Preis       As String * 9
  mw          As String * 1
  bs          As Byte
  datum       As String * 2
  user        As Integer
  LaufNr      As Integer
  gebuehren   As String * 1
  wegmark     As String * 1
  AVP         As String * 1
  zeit        As String * 2 'Integer
  bon         As String * 1
  wg          As String * 2
  knr         As Integer
  gebsumme    As String * 8
  pflag       As Byte
  PersCode    As Byte
  angefordert As String * 1
  FremdGeld   As Byte
  FremdBetrag As String * 8
  TxtTyp      As String * 1
  RezEan      As String * 13
  KKNr        As String * 9
  KKTyp       As Byte
  ZuzaFlag    As String * 1
  rest        As String * 9
  multi       As Byte
End Type

Private Type KZustellType
    knr As String * 5
    Name As String * 30
    Strasse As String * 30
    PlzOrt As String * 30
    telefon As String * 30
    zeit As String * 5              'nicht mehr verwendet außer zum umkopieren in zeitraum
    Geschenke As Integer
    Zahlungsart As String * 1
    KarteNr As String * 19
    gueltig As String * 4
    zHdn As String * 30
    Zeitraum As String * 19
    rest As String * 51
End Type
Dim kzu As KZustellType


Dim af As AnfertStruct
Dim aw As AnfWasStruct
Dim ai(9) As AnfInfoStruct

Dim BesAconto$
Dim BesPzn$
Dim BesMenge%

Dim BelegtKz(2) As String * 48
Dim GedrucktKz(2) As String * 48
Dim afFix As String * 48
Dim awFix As String * 32
Dim BlockNummer&

Dim VkRec As VerkaufStruct

'Dim BesorgerAconto$
'Dim BesorgerMenge%

Dim ANFERT%
Dim ANFBLOCK%

Dim KistenNr%

Public afVonWann$


Private Const DefErrModul = "KISTE.CLS"

Function OpenDatei%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatei%")
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

ret% = False

ANFERT% = clsDat.FileOpen("ANFERT.DAT", "RW", "R", Len(af))
ANFBLOCK% = clsDat.FileOpen("ANFBLOCK.DAT", "RW", "R", Len(aw))
If (ANFERT% > 0) And (ANFBLOCK% > 0) Then
'    Call clsDat.iLock(ANFERT%, 1)
    For i% = 0 To 2
        Get #ANFERT%, i% + 2, BelegtKz$(i%)
    Next i%
    For i% = 0 To 2
        Get #ANFERT%, , GedrucktKz$(i%)
    Next i%
    ret% = True
End If

OpenDatei% = ret%
Call clsError.DefErrPop
End Function

Sub CloseDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CloseDatei")
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
If (ANFERT% > 0) Then
'    Call clsDat.iUnLock(ANFERT%, 1)
    Close #ANFERT%
End If
If (ANFBLOCK% > 0) Then
    Close #ANFBLOCK%
End If

Call clsError.DefErrPop
End Sub

Function Belegt%(nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Belegt%")
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

Belegt% = CheckKistenBit%(0, nr%)

Call clsError.DefErrPop
End Function

Function Gedruckt%(nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Gedruckt%")
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

Gedruckt% = CheckKistenBit%(1, nr%)

Call clsError.DefErrPop
End Function

Function CheckKistenBit%(iModus%, nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckKistenBit%")
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
Dim satz%, satz1%, header%, bit%, BitMaske%
Dim ch$, h$

satz1% = (nr% - 1) \ 8
satz% = satz1% \ 48

If (iModus% = 0) Then
    afFix = BelegtKz$(satz%)
Else
    afFix = GedrucktKz$(satz%)
End If

ch$ = Mid$(afFix, (satz1% Mod 48) + 1, 1)

header% = Asc(ch$)
bit% = 7 - ((nr% - 1) Mod 8)
BitMaske% = 2 ^ bit%

CheckKistenBit% = (header% And BitMaske%)

Call clsError.DefErrPop
End Function

Sub GetKiste(nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetKiste")
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

KistenNr% = nr%
Get #ANFERT%, nr% + 7, af

Call clsError.DefErrPop
End Sub

Sub PutKiste(nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutKiste")
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
Put #ANFERT%, nr% + 7, af
Call clsError.DefErrPop
End Sub

Sub ClearKiste(nr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ClearKiste")
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
Dim i%, bit%, BitMaske%, satz%, pyte%, LaufNr%
Dim ppyte$, ByteLeiste$, sLaufNr$

afFix = String(Len(afFix), 0)
Put #ANFERT%, nr% + 7, afFix

'* belegt-Flag entfernen
bit% = 7 - ((nr% - 1) Mod 8)
BitMaske% = Not (2 ^ bit%) And &HFF
satz% = ((nr% - 1) \ 8) + 1

ByteLeiste$ = ""
For i% = 1 To 3
    Get #ANFERT%, i% + 1, afFix: ByteLeiste$ = ByteLeiste$ + afFix
Next i%

pyte% = Asc(Mid$(ByteLeiste$, satz%, 1))
ppyte$ = Chr$(pyte% And BitMaske%)
Mid$(ByteLeiste$, satz%, 1) = ppyte$
For i% = 1 To 3
    afFix = Mid$(ByteLeiste$, (i% - 1) * Len(af) + 1, Len(af))
    Put #ANFERT%, i% + 1, afFix
Next i%

'* gedruckt-Flag entfernen
ByteLeiste$ = ""
For i% = 1 To 3
    Get #ANFERT%, i% + 4, afFix: ByteLeiste$ = ByteLeiste$ + afFix
Next i%
pyte% = Asc(Mid$(ByteLeiste$, satz%, 1))
ppyte$ = Chr$(pyte% And BitMaske%)
Mid$(ByteLeiste$, satz%, 1) = ppyte$
For i% = 1 To 3
    afFix = Mid$(ByteLeiste$, (i% - 1) * Len(af) + 1, Len(af))
    Put #ANFERT%, i% + 4, afFix
Next i%

Get #ANFERT%, 1, afFix
sLaufNr$ = Mid$(afFix, 3, 1)
LaufNr% = Asc(sLaufNr$) + 1: If (LaufNr% > 127) Then LaufNr% = 0
Mid$(afFix, 3, 1) = Chr$(LaufNr%)
Put #ANFERT%, 1, afFix

Call clsError.DefErrPop
End Sub

Function GetInhalt%(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetInhalt%")
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
Dim h$, h1$, h2$

ret% = False

BesAconto$ = ""
BesPzn$ = ""
BesMenge% = 1

BlockNummer& = af.ANFBLOCK(ind%)
If (BlockNummer& > 0) Then
    Get #ANFBLOCK%, ((BlockNummer& - 1) * 11 + 40) + 1, aw
    For i% = 0 To 9
        Get #ANFBLOCK%, , ai(i%)
    Next i%
    If (InStr("RTHBP", aw.WasTun) > 0) And (aw.Status > 0) Then
        For i% = 0 To 5
            h$ = RTrim$(ai(i%).InfoZeile)
            h1$ = Mid$(h$, 6, 2)
            h2$ = Left$(h$, 7)
            If (h1$ = "A=" Or h1$ = "G=" Or h2$ = "ACONTO=" Or h2$ = "GEBÜHR=") Then
                BesAconto$ = Mid$(h$, 8, 10)
                BesMenge% = Val(Mid$(h$, 21, 4))
'??                If (BesMenge% = 0) Then BesMenge% = 1
'                Exit For
            End If
            If (Mid$(h$, 14, 4) = "PZN=") Then
                BesPzn$ = Mid$(h$, 18, 7)
            End If
        Next i%
        If (aw.WasTun = "R") Then BesPzn$ = "9999999"
    End If
    ret% = True
End If

GetInhalt% = ret%

Call clsError.DefErrPop
End Function

Function PutInhalt%(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutInhalt%")
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

ret% = False

BlockNummer& = af.ANFBLOCK(ind%)
If (BlockNummer& > 0) Then
    Put #ANFBLOCK%, ((BlockNummer& - 1) * 11 + 40) + 1, aw
    For i% = 0 To 9
        Put #ANFBLOCK%, , ai(i%)
    Next i%
    ret% = True
End If

PutInhalt% = ret%

Call clsError.DefErrPop
End Function

Function ClearInhalt%(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ClearInhalt%")
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
Dim i%, ret%, satz%, pyte%, bit%, char%, BitMaske%

ret% = False

BlockNummer& = af.ANFBLOCK(ind%)
If (BlockNummer& > 0) Then
    awFix = String(Len(awFix), 0)
    Put #ANFBLOCK%, ((BlockNummer& - 1) * 11 + 40) + 1, awFix
    For i% = 0 To 9
        Put #ANFBLOCK%, , awFix
    Next i%
    
    satz% = ((BlockNummer& - 1) \ (Len(aw) * 8)) + 1
    pyte% = ((BlockNummer& - 1) Mod (Len(aw) * 8)) \ 8 + 1
    bit% = 7 - ((BlockNummer& - 1) Mod 8)

    Get #ANFBLOCK%, satz%, awFix
    char% = Asc(Mid$(awFix, pyte%, 1))
    BitMaske% = 2 ^ bit%
    char% = char% And (Not BitMaske%)
    Mid$(awFix, pyte%, 1) = Chr$(char%)
    Put #ANFBLOCK%, satz%, awFix
    
    ret% = True
End If

ClearInhalt% = ret%

Call clsError.DefErrPop
End Function


Public Property Get VonWann$()
VonWann$ = af.VonWann
End Property
Public Property Let VonWann(ByVal vNewValue$)
af.VonWann = vNewValue$
End Property

Public Property Get VonWo$()
VonWo$ = af.VonWo
End Property
Public Property Let VonWo(ByVal vNewValue$)
af.VonWo = vNewValue$
End Property

Public Property Get Arbeit$()
Arbeit$ = af.Arbeit
End Property
Public Property Let Arbeit(ByVal vNewValue$)
af.Arbeit = vNewValue$
End Property

Public Property Get KistenStatus%()
KistenStatus% = af.Status
End Property
Public Property Let KistenStatus(ByVal vNewValue%)
af.Status = vNewValue%
End Property

Public Property Get ProfilNr%()
ProfilNr% = af.Profil
End Property
Public Property Let ProfilNr(ByVal vNewValue%)
af.Profil = vNewValue%
End Property



Public Property Get Status%()
Status% = aw.Status
End Property
Public Property Let Status(ByVal vNewValue%)
aw.Status = vNewValue%
End Property

Public Property Get WasTun$()
WasTun$ = aw.WasTun
End Property
Public Property Let WasTun(ByVal vNewValue$)
aw.WasTun = vNewValue$
End Property

Public Property Get RezeptNr%()
RezeptNr% = aw.RezeptNr
End Property
Public Property Let RezeptNr(ByVal vNewValue%)
aw.RezeptNr = vNewValue%
End Property

Public Property Get RezeptEan$()
RezeptEan$ = aw.RezEan
End Property
Public Property Let RezeptEan(ByVal vNewValue$)
aw.RezEan = vNewValue$
End Property

Public Property Get Preis#()
Preis# = Val(aw.Preis)
End Property
Public Property Let Preis(ByVal vNewValue#)
Dim i%, h$
h$ = Right$(Space$(9) + Format(vNewValue#, "0.00"), 9)
For i% = 1 To Len(h$)
    If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
Next i%
aw.Preis = h$
End Property

Public Property Get ErstMagSatz&()
ErstMagSatz& = aw.ErstMagSatz
End Property
Public Property Let ErstMagSatz(ByVal vNewValue&)
aw.ErstMagSatz = vNewValue&
End Property

Public Property Get InfoText$(ByVal ind%)
InfoText$ = ai(ind%).InfoZeile
Call OemToChar(InfoText$, InfoText$)
End Property
Public Property Let InfoText(ByVal ind%, ByVal vNewValue$)
Call CharToOem(vNewValue$, vNewValue$)
ai(ind%).InfoZeile = vNewValue$
End Property

Public Property Get Aconto$()
Aconto$ = BesAconto$
End Property

Public Property Get pzn$()
pzn$ = BesPzn$
End Property
Public Property Let pzn(ByVal vNewValue$)
Dim i%
Dim h$

For i% = 0 To 5
    h$ = RTrim$(ai(i%).InfoZeile)
    If (Mid$(h$, 14, 4) = "PZN=") Then
        Mid$(ai(i%).InfoZeile, 18, 7) = Left$(vNewValue$, 7)
        h$ = Mid$(vNewValue$, 8)
        If (h$ = "") Then h$ = Space$(36)
        ai(0).InfoZeile = Left$(h$, 24)
        ai(1).InfoZeile = Mid$(h$, 25) + Mid$(ai(1).InfoZeile, 13)
        Exit For
    End If
Next i%
End Property

Public Property Get menge%()
menge% = BesMenge%
End Property
Public Property Let menge(ByVal vNewValue%)
Dim i%
Dim h$, h1$, h2$

For i% = 0 To 5
    h$ = RTrim$(ai(i%).InfoZeile)
    h1$ = Mid$(h$, 6, 2)
    h2$ = Left$(h$, 7)
    If (h1$ = "A=" Or h1$ = "G=" Or h2$ = "ACONTO=" Or h2$ = "GEBÜHR=") Then
        Mid$(ai(i%).InfoZeile, 21, 4) = Right$(Space$(4) + Format(vNewValue%, "0"), 4)
        Exit For
    End If
Next i%
End Property

Public Property Get HatLieferSchein() As Byte
HatLieferSchein = af.HatLieferSchein
End Property
Public Property Let HatLieferSchein(ByVal vNewValue As Byte)
af.HatLieferSchein = vNewValue
End Property

Public Property Get LieferScheinNr%()
LieferScheinNr% = af.LSnr
End Property
Public Property Let LieferScheinNr(ByVal vNewValue%)
af.LSnr = vNewValue%
End Property



Function PznInKiste%(nr%, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PznInKiste%")
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
Dim i%, j%, PznGefunden%, ret%
Dim BlockNummer&
Dim aiPzn$, h$

ret% = -1
    
Call GetKiste(nr%)
    
PznGefunden% = False
For i% = 0 To 9
    BlockNummer& = af.ANFBLOCK(i%)
    If (BlockNummer& > 0) Then
'        afVonWann$ = clsAnfert1.Wert("VONWANN")
        Call GetInhalt(i%)
        If (InStr("RTHBP", aw.WasTun) > 0) And (aw.Status > 0) Then
            For j% = 0 To 4
                h$ = RTrim(InfoText$(j%))
                If (Mid$(h$, 14, 4) = "PZN=") Then
                    aiPzn$ = Mid$(h$, 18, 7)
                    If (aiPzn$ = pzn$) Then
                        PznGefunden% = True
                        Exit For
                    End If
                End If
            Next j%
        End If
        If (PznGefunden%) Then Exit For
    End If
Next i%

If (PznGefunden%) Then
    ret% = i%
End If
    
PznInKiste% = ret%

Call clsError.DefErrPop
End Function

Function MachAuftrag%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachAuftrag%")
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
Dim clsKunden1 As New clsKunden
Dim clsKundenZus1 As New clsKundZusatz
Dim AufDB As Database
Dim ProfDB As Database
Dim auf As Recordset
Dim auf2 As Recordset
Dim prof As Recordset
Dim arch As Recordset
Dim i%, j%, KZUSTELL%, KZUST_IDX%, zahlart%, ok%, IstMischKunde%, AuftragAdd%
Dim dBesAconto#, dBesPreis#
Dim recno&
Dim s$, AufNr$, ZustName$(4), Zeitraum$, SQLStr$, LSnr$
Dim tRec2 As Recordset


MachAuftrag% = 0

On Error Resume Next
Err.Clear
Set ProfDB = OpenDatabase("PROFILE.MDB")
Set prof = ProfDB.OpenRecordset("Profile", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    Call clsError.DefErrPop: Exit Function
End If

prof.index = "Unique"
prof.Seek "=", af.Profil
ok% = 0
If (prof.NoMatch = False) Then
    If (prof!AbhFertig) Then
        ok% = True
    End If
End If
If (ok% = 0) Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If

        
dBesAconto# = clsOpTool.FNX(clsOpTool.xVal(BesAconto$))
dBesPreis# = clsOpTool.FNX(BesMenge% * clsOpTool.xVal(aw.Preis))

ok% = True
If (aw.RezeptNr = 0) Then
    'Preisunterschied kann dem Kunden nur verrechnet werden, wenn kein Lieferschein und er jetzt einen Auftrag bekommt
    If (dBesPreis# <> dBesAconto#) And ((prof!AutoZustellung = 0) Or (af.HatLieferSchein)) Then
        ok% = 0
    End If
Else
    'Wenn RezGebühr noch nicht bezahlt ist und bei der Bestellung kein Lieferschien erstellt wurde,
    'kann sie (die RezGebühr)nur im jetzt entstehenden Auftrag verrechnet werden
    If (dBesAconto# > 0#) And (prof!AutoZustellung = 0) And (af.HatLieferSchein = 0) Then
        ok% = 0
    End If
End If

If (ok% = 0) Then
    Call clsDialog.MessageBox("Abholer " + Format(KistenNr%, "0") + " muß an der Kasse fertiggemacht werden!", vbInformation)
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If


Err.Clear
Set AufDB = OpenDatabase("AUFTRAG.MDB")
Set auf = AufDB.OpenRecordset("Aufträge", dbOpenTable)
If Err.Number <> 0 Then
    On Error GoTo DefErr
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If
 

On Error GoTo DefErr
KZUSTELL% = clsDat.FileOpen("kzustell.dat", "R", "R", 256)
If (KZUSTELL% = 0) Then
    AufDB.Close
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If
 
FabsErrf% = clsDat.IndexOpen("kzustell.ix1", KZUST_IDX%)
If (FabsErrf%) Then
    Close #KZUSTELL%
    AufDB.Close
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If

MachAuftrag% = True

If (prof!AutoZustellung = 0) Then
    Close #KZUSTELL%
    AufDB.Close
    ProfDB.Close
    Call clsError.DefErrPop: Exit Function
End If

clsKundenZus1.OpenDatei
clsKundenZus1.GetRecord (af.KundenNr + 1)
IstMischKunde% = clsKundenZus1.MischKunde
clsKundenZus1.CloseDatei
Set clsKundenZus1 = Nothing

AuftragAdd% = 0
If (IstMischKunde% = 0) Then
    SQLStr$ = "SELECT * FROM AUFTRÄGE WHERE KNR = " + Str$(af.KundenNr)
    SQLStr$ = SQLStr$ + " AND ISNULL(LIEFERLISTE)"
    SQLStr$ = SQLStr$ + " AND FERTIG=0"
    Set auf2 = AufDB.OpenRecordset(SQLStr$)
    If (auf2.EOF = False) Then
        auf2.MoveFirst
        Do
            If (auf2.EOF) Then
                Exit Do
            End If
            
            If (DateDiff("s", auf2!datum, Now) < 60) Then
                AuftragAdd% = True
                Exit Do
            End If
            
            auf2.MoveNext
        Loop
    End If
End If
    
If (AuftragAdd%) Then
    Close #KZUSTELL%
    FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
    
    auf2.Edit
        
    If (aw.RezeptNr > 0) Then
        auf2!RezAnz = 1  'aw.RezeptNr
        If (af.HatLieferSchein) Then
            If (af.LSnr > 0) Then
                If (Val(BesPzn$) > 0) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
                    Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
                    If (tRec2.EOF) Then
                    Else
                        auf2!Summe = auf2!Summe + tRec2!ZuzaNeu / 100
                    End If
                End If
            End If
        Else
            auf2!Summe = auf2!Summe + Val(BesAconto$)
        End If
    Else
        If (af.HatLieferSchein) Then
            If (af.LSnr > 0) Then
                auf2!Summe = auf2!Summe + Val(aw.Preis)
            Else
            End If
        Else
            auf2!Summe = auf2!Summe + Val(aw.Preis) - Val(BesAconto$)
        End If
    End If
    
    auf2!Positionen = auf2!Positionen + 1 'ZustellAnz
    
    Err.Clear
    auf2.Update
Else
    auf.index = "Unique"
    AufNr$ = EAN_AUFTRAG + Format(Now, "YYMMDD") + "00010"
    s$ = Space(15)
    i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, INIFILE_DP)
    If (i% > 0) Then
        s$ = Left(s$, i%)
        If (Left(s$, 2) = EAN_AUFTRAG) And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
            AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
        End If
    End If
        
    Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
    Do
        i% = 0
        Call clsOpTool.EanPruef(AufNr$)
        auf.Seek "=", AufNr$
        If Not (auf.NoMatch) Then i% = 1
        arch.index = "Unique"
        arch.Seek "=", AufNr$
        If Not (arch.NoMatch) Then i% = 1
        If (i% = 1) Then
            AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
        End If
    Loop While (i% = 1)
    i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, INIFILE_DP)
        
    zahlart% = 1
    FabsErrf% = clsDat.IndexSearch%(KZUST_IDX%, Format(af.KundenNr, "00000"), recno&)
    If (recno& > 0) Then
        Get #KZUSTELL%, recno& + 1, kzu
        
        zahlart% = Val(kzu.Zahlungsart)
        ZustName$(0) = Trim(kzu.Name)
        ZustName$(1) = Trim(kzu.zHdn)
        ZustName$(2) = Trim(kzu.Strasse)
        ZustName$(3) = Trim(kzu.PlzOrt)
        ZustName$(4) = Trim(kzu.telefon)
        Zeitraum$ = Trim(kzu.Zeitraum)
    End If
    
    Close #KZUSTELL%
    FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
    
    
    clsKunden1.OpenDatei
    clsKunden1.GetRecord (af.KundenNr + 1)
    
    If (ZustName$(0) = "") And (ZustName$(1) = "") Then
        ZustName$(0) = RTrim$(clsKunden1.Wert("name", 0))
        ZustName$(1) = RTrim$(clsKunden1.Wert("name", 1))
    End If
    If (ZustName$(2) = "") And (ZustName$(3) = "") Then
        ZustName$(2) = RTrim$(clsKunden1.Wert("name", 2))
        ZustName$(3) = RTrim$(clsKunden1.Wert("name", 3))
    End If
    'If (zuName(4) = "") Then
    '    GetDataRecord dbKZus, AktKundensatz
    '    zuName(4) = Trim(kzus.Tel)
    'End If
    For i% = 0 To 3
        Call OemToChar(ZustName$(i%), ZustName$(i%))
    Next i%
    Call OemToChar(Zeitraum$, Zeitraum$)
        
    auf.AddNew
    auf!AuftragsNr = AufNr$
    auf!ProfilNr = af.Profil
    auf!knr = af.KundenNr
    auf!datum = Now
    
    LSnr$ = ""
    If (af.LSnr > 0) Then
        LSnr$ = Format(Now, "YY") + "/" + Format(af.LSnr, "00000") + "V"
    End If
    auf!LSnr = LSnr$
        
    auf!Zahlungsart = zahlart%
    If (aw.RezeptNr > 0) Then
        auf!RezAnz = 1  'aw.RezeptNr
        If (af.HatLieferSchein) Then
            If (af.LSnr > 0) Then
                If (Val(BesPzn$) > 0) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(BesPzn$)
                    Set tRec2 = TaxeDB.OpenRecordset(SQLStr$)
                    If (tRec2.EOF) Then
                        auf!Summe = 0
                    Else
                        auf!Summe = tRec2!ZuzaNeu / 100
                    End If
                Else
                    auf!Summe = 0
                End If
            Else
                auf!Summe = 0
            End If
        Else
            auf!Summe = Val(BesAconto$)
        End If
    Else
        auf!RezAnz = 0
        If (af.HatLieferSchein) Then
            If (af.LSnr > 0) Then
                auf!Summe = Val(aw.Preis)
            Else
                auf!Summe = 0
            End If
        Else
            auf!Summe = Val(aw.Preis) - Val(BesAconto$)
        End If
    End If
    
    prof.index = "Unique"
    prof.Seek "=", auf!ProfilNr
    If (prof.NoMatch) Then
        auf!RezBringen = prof!RezBringen
    Else
        auf!RezBringen = 0
    End If
    
    auf!Positionen = 1  'ZustellAnz
    auf!Kundenichtda = False
    auf!kZu1 = ZustName(0)
    auf!kZu2 = ZustName(1)
    auf!kZu3 = ZustName(2)
    auf!kZu4 = ZustName(3)
    auf!kZu5 = ZustName(4)
    auf!Zeitraum = Zeitraum$
    
    Err.Clear
    auf.Update

    clsKunden1.CloseDatei
    Set clsKunden1 = Nothing
End If
    
AufDB.Close
ProfDB.Close


s$ = Space(10)
i% = GetPrivateProfileString("LIFMANAG", "Aktualisieren", "", s$, 255, "winop.ini")
If (i% > 0) Then
    i% = Val(Left(s$, i%))
Else
    i% = 0
End If
If (i% >= 32000) Then
    i% = 0
End If
i% = i% + 1
s$ = Format(i%, "0")
j% = WritePrivateProfileString("LIFMANAG", "Aktualisieren", s$, "winop.ini")

Call clsError.DefErrPop
End Function

'Function MachAuftrag%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("MachAuftrag%")
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
'Dim AufDB As Database
'Dim ProfDB As Database
'Dim auf As Recordset
'Dim prof As Recordset
'Dim arch As Recordset
'Dim i%, j%, KZUSTELL%, KZUST_IDX%, zahlart%, ok%, IstNeu%
'Dim dBesAconto#, dBesPreis#
'Dim recno&
'Dim s$, AufNr$, ZustName$(4), Zeitraum$
'
'
'MachAuftrag% = 0
'
'On Error Resume Next
'Err.Clear
'Set ProfDB = OpenDatabase("PROFILE.MDB")
'Set prof = ProfDB.OpenRecordset("Profile", dbOpenTable)
'If Err.Number <> 0 Then
'    On Error GoTo DefErr
'    Call clsError.DefErrPop: Exit Function
'End If
'
'prof.index = "Unique"
'prof.Seek "=", af.Profil
'ok% = 0
'If (prof.NoMatch = False) Then
'    If (prof!AbhFertig) Then
'        ok% = True
'    End If
'End If
'If (ok% = 0) Then
'    On Error GoTo DefErr
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'dBesAconto# = clsOpTool.FNX(clsOpTool.xVal(BesAconto$))
'dBesPreis# = clsOpTool.FNX(BesMenge% * clsOpTool.xVal(aw.Preis))
'
'ok% = True
'If (aw.RezeptNr = 0) Then
'    'Preisunterschied kann dem Kunden nur verrechnet werden, wenn kein Lieferschein und er jetzt einen Auftrag bekommt
'    If (dBesPreis# <> dBesAconto#) And ((prof!AutoZustellung = 0) Or (af.HatLieferSchein)) Then
'        ok% = 0
'    End If
'Else
'    'Wenn RezGebühr noch nicht bezahlt ist und bei der Bestellung kein Lieferschien erstellt wurde,
'    'kann sie (die RezGebühr)nur im jetzt entstehenden Auftrag verrechnet werden
'    If (dBesAconto# > 0#) And (prof!AutoZustellung = 0) And (af.HatLieferSchein = 0) Then
'        ok% = 0
'    End If
'End If
'
'If (ok% = 0) Then
'    Call MsgBox("Abholer " + Format(KistenNr%, "0") + " muß an der Kasse fertiggemacht werden!", vbInformation)
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'Err.Clear
'Set AufDB = OpenDatabase("AUFTRAG.MDB")
'Set auf = AufDB.OpenRecordset("Aufträge", dbOpenTable)
'If Err.Number <> 0 Then
'    On Error GoTo DefErr
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'On Error GoTo DefErr
'KZUSTELL% = clsDat.FileOpen("kzustell.dat", "R", "R", 256)
'If (KZUSTELL% = 0) Then
'    AufDB.Close
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'FabsErrf% = clsDat.IndexOpen("kzustell.ix1", KZUST_IDX%)
'If (FabsErrf%) Then
'    Close #KZUSTELL%
'    AufDB.Close
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'MachAuftrag% = True
'
'If (prof!AutoZustellung = 0) Then
'    Close #KZUSTELL%
'    AufDB.Close
'    ProfDB.Close
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'auf.index = "Unique"
'AufNr$ = EAN_AUFTRAG + Format(Now, "YYMMDD") + "00010"
's$ = Space(15)
'i% = GetPrivateProfileString("Allgemein", "DSKAuftragsNr", "", s$, 255, INIFILE_DP)
'If (i% > 0) Then
'    s$ = Left(s$, i%)
'    If (Left(s$, 2) = EAN_AUFTRAG) And (Mid(s$, 3, 6) = Format(Now, "YYMMDD")) Then
'        AufNr$ = Left(s$, 8) + Format(Val(Mid(s$, 9, 4)) + 1, "0000") + "0"
'    End If
'End If
'
'Set arch = AufDB.OpenRecordset("Archiv", dbOpenTable)
'Do
'    i% = 0
'    Call clsOpTool.EanPruef(AufNr$)
'    auf.Seek "=", AufNr$
'    If Not (auf.NoMatch) Then i% = 1
'    arch.index = "Unique"
'    arch.Seek "=", AufNr$
'    If Not (arch.NoMatch) Then i% = 1
'    If (i% = 1) Then
'        AufNr$ = Left(AufNr$, 8) + Format(Val(Mid(AufNr$, 9, 4)) + 1, "0000") + "0"
'    End If
'Loop While (i% = 1)
'i% = WritePrivateProfileString("Allgemein", "DSKAuftragsNr", AufNr$, INIFILE_DP)
'
'zahlart% = 1
'FabsErrf% = clsDat.IndexSearch%(KZUST_IDX%, Format(af.KundenNr, "00000"), recno&)
'If (recno& > 0) Then
'    Get #KZUSTELL%, recno& + 1, kzu
'
'    zahlart% = Val(kzu.Zahlungsart)
'    ZustName$(0) = Trim(kzu.Name)
'    ZustName$(1) = Trim(kzu.zHdn)
'    ZustName$(2) = Trim(kzu.Strasse)
'    ZustName$(3) = Trim(kzu.PlzOrt)
'    ZustName$(4) = Trim(kzu.telefon)
'    Zeitraum$ = Trim(kzu.Zeitraum)
'End If
'
'Close #KZUSTELL%
'FabsErrf% = clsDat.IndexClose(KZUST_IDX%)
'
'
'On Error Resume Next
's$ = Kun1.Beleg
'IstNeu% = (Err = 91)
'On Error GoTo DefErr
'
'If (IstNeu%) Then
'    Set Kun1 = New clsKunden
'    Kun1.OpenDatei
'End If
'
'Kun1.GetRecord (af.KundenNr + 1)
'If (ZustName$(0) = "") And (ZustName$(1) = "") Then
'    ZustName$(0) = RTrim$(Kun1.Name(0))
'    ZustName$(1) = RTrim$(Kun1.Name(1))
'End If
'If (ZustName$(2) = "") And (ZustName$(3) = "") Then
'    ZustName$(2) = RTrim$(Kun1.Name(2))
'    ZustName$(3) = RTrim$(Kun1.Name(3))
'End If
'
'If (IstNeu%) Then
'    Kun1.CloseDatei
'    Set Kun1 = Nothing
'End If
'
'
''clsKunden1.OpenDatei
''clsKunden1.GetRecord (af.KundenNr + 1)
''
''If (ZustName$(0) = "") And (ZustName$(1) = "") Then
''    ZustName$(0) = RTrim$(clsKunden1.Wert("name", 0))
''    ZustName$(1) = RTrim$(clsKunden1.Wert("name", 1))
''End If
''If (ZustName$(2) = "") And (ZustName$(3) = "") Then
''    ZustName$(2) = RTrim$(clsKunden1.Wert("name", 2))
''    ZustName$(3) = RTrim$(clsKunden1.Wert("name", 3))
''End If
'
''If (zuName(4) = "") Then
''    GetDataRecord dbKZus, AktKundensatz
''    zuName(4) = Trim(kzus.Tel)
''End If
'For i% = 0 To 3
'    Call OemToChar(ZustName$(i%), ZustName$(i%))
'Next i%
'Call OemToChar(Zeitraum$, Zeitraum$)
'
'auf.AddNew
'auf!AuftragsNr = AufNr$
'auf!ProfilNr = af.Profil
'auf!knr = af.KundenNr
'auf!datum = Now
'auf!LSNr = ""
'
'auf!Zahlungsart = zahlart%
'If (aw.RezeptNr > 0) Then
'    auf!RezAnz = 1  'aw.RezeptNr
'    If (af.HatLieferSchein) Then
'        auf!Summe = 0
'    Else
'        auf!Summe = Val(BesAconto$)
'    End If
'Else
'    auf!RezAnz = 0
'    If (af.HatLieferSchein) Then
'        auf!Summe = 0
'    Else
'        auf!Summe = Val(aw.Preis) - Val(BesAconto$)
'    End If
'End If
'
'prof.index = "Unique"
'prof.Seek "=", auf!ProfilNr
'If (prof.NoMatch) Then
'    auf!RezBringen = prof!RezBringen
'Else
'    auf!RezBringen = 0
'End If
'
'auf!Positionen = 1  'ZustellAnz
'auf!Kundenichtda = False
'auf!kZu1 = ZustName(0)
'auf!kZu2 = ZustName(1)
'auf!kZu3 = ZustName(2)
'auf!kZu4 = ZustName(3)
'auf!kZu5 = ZustName(4)
'auf!Zeitraum = Zeitraum$
'
'Err.Clear
'auf.Update
'
'AufDB.Close
'ProfDB.Close
'
''clsKunden1.CloseDatei
''Set clsKunden1 = Nothing
'
's$ = Space(10)
'i% = GetPrivateProfileString("LIFMANAG", "Aktualisieren", "", s$, 255, "winop.ini")
'If (i% > 0) Then
'    i% = Val(Left(s$, i%))
'Else
'    i% = 0
'End If
'If (i% >= 32000) Then
'    i% = 0
'End If
'i% = i% + 1
's$ = Format(i%, "0")
'j% = WritePrivateProfileString("LIFMANAG", "Aktualisieren", s$, "winop.ini")
'
'Call clsError.DefErrPop
'End Function

Sub KistenReorg()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("KistenReorg")
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
Dim i%, j%, KistenNr%, ANFERT_NEW%, ANFBLOCK_NEW%, AnzBlock%, TmpRec%, KistenOk%
Dim satz%, pyte%, bit%, char%, BitMaske%
Dim BlockNummerNeu&

Call OpenDatei
        
If (Dir$("ANFERT.NEW") <> "") Then
    Kill "ANFERT.NEW"
End If
ANFERT_NEW% = clsDat.FileOpen("ANFERT.NEW", "RW", "R", Len(af))

Get #ANFERT%, 1, af
Put #ANFERT_NEW%, 1, af

'BelegtKz
For i% = 1 To 3
    Get #ANFERT%, i% + 1, af
    Put #ANFERT_NEW%, i% + 1, af
Next i%

'DruckKz
For i% = 4 To 6
    Get #ANFERT%, i% + 1, af
    Put #ANFERT_NEW%, i% + 1, af
Next i%


If (Dir$("ANFBLOCK.NEW") <> "") Then
    Kill "ANFBLOCK.NEW"
End If
ANFBLOCK_NEW% = clsDat.FileOpen("ANFBLOCK.NEW", "RW", "R", Len(aw))

awFix = String(Len(awFix), 0)
For i% = 1 To 40
    Put #ANFBLOCK_NEW%, i%, awFix
Next i%

        
For KistenNr% = 1 To 999
    If (KistenNr = 975) Then
        Beep
    End If
    
    KistenOk = 0
    If (Belegt(KistenNr%)) Then
        Call GetKiste(KistenNr%)

        For i% = 0 To 9
            BlockNummer& = af.ANFBLOCK(i%)
            If (BlockNummer& > 0) Then
                Get #ANFBLOCK%, ((BlockNummer& - 1) * 11 + 40) + 1, aw
                If (InStr("RTHBP", aw.WasTun) = 0) Or (aw.Status = 0) Then
'                    Debug.Print Str$(KistenNr%)
                    BlockNummer& = 0
                End If
            End If
                
            If (BlockNummer& > 0) Then
                KistenOk% = True
'                If (aw.ErstMagSatz > 0) Then
'                    Debug.Print Str$(KistenNr%)
'                End If
                
                For j% = 0 To 9
                    Get #ANFBLOCK%, , ai(j%)
                Next j%
                
                BlockNummerNeu& = BlockNummerNeu& + 1
                
                Put #ANFBLOCK_NEW%, ((BlockNummerNeu& - 1) * 11 + 40) + 1, aw
                For j% = 0 To 9
                    Put #ANFBLOCK_NEW%, , ai(j%)
                Next j%
                
                'BelegtKz in anfblock setzen
                satz% = ((BlockNummerNeu& - 1) \ (Len(aw) * 8)) + 1
                pyte% = ((BlockNummerNeu& - 1) Mod (Len(aw) * 8)) \ 8 + 1
                bit% = 7 - ((BlockNummerNeu& - 1) Mod 8)
            
                Get #ANFBLOCK_NEW%, satz%, awFix
                char% = Asc(Mid$(awFix, pyte%, 1))
                BitMaske% = 2 ^ bit%
                char% = char% Or BitMaske%  'And (Not bitmaske%)
                Mid$(awFix, pyte%, 1) = Chr$(char%)
                Put #ANFBLOCK_NEW%, satz%, awFix
                                        
                af.ANFBLOCK(i%) = BlockNummerNeu&
            Else
                af.ANFBLOCK(i%) = 0
            End If
        Next i%
    End If
        
    If (KistenOk) Then
        Put #ANFERT_NEW%, KistenNr% + 7, af
    Else
        afFix = String(Len(afFix), 0)
        Put #ANFERT_NEW%, KistenNr% + 7, afFix
    End If
Next KistenNr%

        
Call CloseDatei
Close #ANFERT_NEW%
Close #ANFBLOCK_NEW%

If (Dir$("ANFERT.OLD") <> "") Then
    Kill "ANFERT.OLD"
End If
Name "ANFERT.DAT" As "ANFERT.OLD"

If (Dir$("ANFBLOCK.OLD") <> "") Then
    Kill "ANFBLOCK.OLD"
End If
Name "ANFBLOCK.DAT" As "ANFBLOCK.OLD"
        
Name "ANFERT.NEW" As "ANFERT.DAT"
Name "ANFBLOCK.NEW" As "ANFBLOCK.DAT"

Call clsError.DefErrPop
End Sub

Sub SpeicherDosBezug(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherDosBezug")
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
Dim i%, j%, iFixiert%, iZeit%
Dim fAnzPackungen!, MaxBezugRec!
Dim FabsRecno&
Dim dpreis#
Dim pzn$, key$, txt$, RezEan$, SQLStr$
Dim KundBez1 As New clsKunBezug

'Kiste1.GetInhalt (ind%)
If (aw.RezeptNr > 0) And (Val(BesAconto) = 0) Then
    Call clsError.DefErrPop: Exit Sub
End If
iZeit% = Val(Format(Now, "HHMM"))

KundBez1.OpenDatei ("RW")

FabsErrf% = 0
With KundBez1
    .GetFirstRecord
    MaxBezugRec! = .erstmax

    pzn$ = BesPzn
    fAnzPackungen! = BesMenge
    
    RezEan$ = clsOpTool.Bcd2ascii(aw.RezEan, 12)
    
    If (aw.RezeptNr > 0) Then   'Kassenrezept
        .pzn = "9999999"
        .Preis = MakeBezugPreis(Val(BesAconto))
        txt = "OFFENE GEBÜHREN " + RezEan
         .multi = 1
    Else    'Privatrezept
        .pzn = BesPzn
        .Preis = MakeBezugPreis(Val(aw.Preis))
        
        txt$ = InfoText(0)
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            Call clsOpTool.Taxe2ast(pzn$)
            txt$ = Ast1.kurz + Ast1.meng + " " + Ast1.meh
        End If
        .multi = BesMenge
    End If
    
    Call CharToOem(txt$, txt$)
    .text = txt$
        
    .mw = "0"
    .wg = "10"
    .user = Val(Para1.user)
    
    .wegmark = " "
    .bon = " "
    
    .zeit = iZeit%
    
    .RezEan = RezEan
        
    .knr = af.KundenNr
    .datum = clsOpTool.iDate(Format(Now, "DDMMYY"))
    .KnrIdx = Right$(Space$(5) + Str$(af.KundenNr), 5)
    .DatIdx = Format(Now, "YYYYMMDD")
    .BelegIdx = Left$("W" + Right(String(4, "0") + Mid(Str(Val(Format(Now, "HHMM"))), 2), 4) + Right(Format(Now, "SS"), 1) + Space$(10), Len(.BelegIdx))
    .LfdIdx = Right$(Space$(3) + Str$(1), 3)
    
'    .Absender = 0
'    .GesendetAn = 0
    
    key$ = .KnrIdx + .DatIdx + .BelegIdx + .LfdIdx
    FabsErrf% = .IndexInsert(0, key$, FabsRecno&)
'    If (FabsErrf% <> 0) Then
'        Exit For
'    End If
    
    If (MaxBezugRec! < FabsRecno&) Then
        MaxBezugRec! = FabsRecno&
    End If
    .PutRecord (FabsRecno& + 1)
    
    .GetFirstRecord
    .erstmax = MaxBezugRec!
    .PutRecord (1)
    
    If (FabsErrf% <> 0) Then
        Call clsDialog.MessageBox("Problem:" + vbCrLf + vbCrLf + "Fehler" + Str$(FabsErrf%) + " in der Bezugsdatei!" + vbCrLf, vbCritical)
    End If
End With

KundBez1.CloseDatei
Set KundBez1 = Nothing

Call clsError.DefErrPop
End Sub

Function MakeBezugPreis$(Preis#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeBezugPreis$")
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
Dim x%
Dim ret$

ret$ = Str$(Preis#)
x% = InStr(ret$, ".")
'* maximal 2 Nachkommata
If (x% > 0) Then
    ret$ = Mid$(ret$, 1, x% + 2)
End If
If (x% = 0) Then
    ret$ = ret$ + ".00"
ElseIf (x% = Len(ret$) - 1) Then
    ret$ = ret$ + "0"
End If

MakeBezugPreis$ = ret$

Call clsError.DefErrPop
End Function


