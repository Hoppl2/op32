Attribute VB_Name = "modMatch"
Option Explicit

Public Const ANZMATCHKONVERTIERUNGEN = 24

Public Const TAXE_ALPHA = 0
Public Const TAXE_MATCH = 1
Public Const LAGER_MATCH = 2

Public TaxeAuswahlModus%
Public MatchModus%

Public MatchAnzeigeTyp$(2)

Public OrgSuch$

Dim MatchKlein$, MatchGross$, MatchGueltig$
Dim MatchPhonAlt$(ANZMATCHKONVERTIERUNGEN)
Dim MatchPhonNeu$(ANZMATCHKONVERTIERUNGEN)

Private Const DefErrModul = "winmatch.bas"

