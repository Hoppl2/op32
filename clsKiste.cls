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

Private Type AnfertStruct
    Status As String * 1
    VonWann As String * 4
    VonWo As Byte 'String * 1
    Mitarbeiter As String * 1
    arbeit As String * 5
    Angef As String * 1
    AngefWann As String * 4
    AngefVonWo As String * 1
    ANFBLOCK(0 To 9) As Integer
    KundenNr As Integer
    rest As String * 8
End Type

Private Type AnfWasStruct
    Status As String * 1
    WasTun As String * 1
    RezeptNr As Integer
    preis As String * 9
    ErstMagSatz As Long
    aconto As String * 9
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
  preis       As String * 9
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
  pFlag       As Byte
  PersCode    As Byte
  Angefordert As String * 1
  FremdGeld   As Byte
  FremdBetrag As String * 8
  TxtTyp      As String * 1
  RezEan      As String * 13
  KKNr        As String * 9
  KKTyp       As Byte
  ZuzaFlag    As String * 1
  rest        As String * 9
  Multi       As Byte
End Type


Dim af As AnfertStruct
Dim aw As AnfWasStruct
Dim ai(9) As AnfInfoStruct

Dim BelegtKz(2) As String * 48
Dim afFix As String * 48
Dim BlockNummer&

Dim VkRec As VerkaufStruct

Dim BesorgerAconto$
Dim BesorgerMenge%

Dim ANFERT%
Dim ANFBLOCK%

Public afVonWann$

Private Const DefErrModul = "wKiste.bas"

Function OpenDatei%()
Dim i%, ret%

ret% = False

ANFERT% = clsDat.FileOpen("ANFERT.DAT", "R", "R", Len(af))
ANFBLOCK% = clsDat.FileOpen("ANFBLOCK.DAT", "R", "R", Len(aw))
If (ANFERT% > 0) And (ANFBLOCK% > 0) Then
    Call clsDat.iLock(ANFERT%, 1)
    For i% = 0 To 2
        Get #ANFERT%, i% + 2, BelegtKz$(i%)
    Next i%
    ret% = True
End If

OpenDatei% = ret%
End Function

Sub CloseDatei()

If (ANFERT% > 0) Then
    Call clsDat.iUnLock(ANFERT%, 1)
    Close #ANFERT%
End If
If (ANFBLOCK% > 0) Then
    Close #ANFBLOCK%
End If

End Sub

Function Belegt%(nr%)
Dim satz%, satz1%, header%, Bit%, bitmaske%
Dim ch$, h$

satz1% = (nr% - 1) \ 8
satz% = satz1% \ 48
'Get #ANFERT%, satz% + 2, afFix
afFix = BelegtKz$(satz%)
ch$ = Mid$(afFix, (satz1% Mod 48) + 1, 1)

header% = Asc(ch$)
Bit% = 7 - ((nr% - 1) Mod 8)
bitmaske% = 2 ^ Bit%

Belegt% = (header% And bitmaske%)

End Function

Sub GetKiste(nr%)
Get #ANFERT%, nr% + 7, af
End Sub

Function GetInhalt%(ind%)
Dim i%, ret%

ret% = False

BlockNummer& = af.ANFBLOCK(ind%)
If (BlockNummer& > 0) Then
    Get #ANFBLOCK%, ((BlockNummer& - 1) * 11 + 40) + 1, aw
    For i% = 0 To 9
        Get #ANFBLOCK%, , ai(i%)
    Next i%
    ret% = True
End If

GetInhalt% = ret%

End Function

Function VonWann$()
VonWann$ = af.VonWann
End Function

Function VonWo$()
VonWo$ = af.VonWo
End Function

Function Status$()
Status$ = aw.Status
End Function

Function WasTun$()
WasTun$ = aw.WasTun
End Function

Function RezeptNr%()
RezeptNr% = aw.RezeptNr
End Function

Function RezeptEan$()
RezeptEan$ = aw.RezEan
End Function

Function InfoText$(ind%)
InfoText$ = ai(ind%).InfoZeile
End Function

Function PznInKiste%(nr%, pzn$)
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
        If (InStr("RTHBP", aw.WasTun) > 0) And (Asc(aw.Status) > 0) Then
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

End Function

