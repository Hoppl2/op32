Attribute VB_Name = "modwmatch"
Option Explicit

Type AusgabeStruct
    Name As String * 200 '100 '80
    DSnummer As Long
    Verweis As Long
    LagerKz As Byte
    pzn As Long
    Bookmark As Variant
    key As String * 38  '10
    ZusatzTextKz As Byte
    BesorgtKz As Byte
End Type


Public Const MATCH_ARTIKEL = 0
Public Const MATCH_LIEFERANTEN = 1
Public Const MATCH_HILFSTAXE = 2

Public Const MATCH_KUNDEN = 55

Public Const TAXE_ALPHA = 0
Public Const TAXE_MATCH = 1
Public Const LAGER_MATCH = 2
Public Const MERKZETTEL_MATCH = 3

Public Const LIEFERANTEN_ALPHA = 0
Public Const LIEFERANTEN_NUMERISCH = 1


Public Match1 As Object
Public Match2 As Object

'Public MatchModus%
Public MatchTyp%
'Public MatchListeTyp%
'Public MatchAnzeigeTyp$()

'Public MatchAutoRet%
'Public MatchNurEiner%

'Public OrgSuch$

Public Ausgabe() As AusgabeStruct

Public AnzAnzeige%
Public MaxAnzeige%

Public buf$
Public KeinRowColChange%

'Public MatchcodePzn$
'Public MatchcodeTxt$
'Public MatchcodeErg$

Public EditModus%
Public EditErg%
Public EditTxt$
Public EditAnzGefunden%
Public EditGef%(19)

Public StammKeyLen%

'Dim MatchKlein$, MatchGross$, MatchGueltig$
'Dim MatchPhonAlt$(ANZMATCHKONVERTIERUNGEN)
'Dim MatchPhonNeu$(ANZMATCHKONVERTIERUNGEN)


Public Const SONDERPR_DB = "SONDERPR.MDB"

Public SonderPrDB As Database
Public SonderPrRec As Recordset
Public SonderPrOk%
Public SonderPrDa%


