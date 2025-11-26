VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMatchLieferanten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const INI_SECTION = "Auswahl-Lieferanten"
Const INFO_SECTION = "Infobereich Auswahl-Lieferanten"

Dim LiefNr%
Dim FabsRecno&

Private Const DefErrModul = "clsMatchLieferanten.cls"

Sub InitMatch()
Dim i%, j%
Dim h$

ReDim MatchAnzeigeTyp$(0)
MatchAnzeigeTyp$(0) = "Lieferanten"

With frmMatchcode.lstSortierung
    .Clear
    j% = 0
    For i% = 1 To 200
        Call Lif1.GetRecord(i% + 1)
        h$ = Lif1.kurz
        h$ = UCase(Trim$(h$))
        If (h$ <> "") Then
            If (Asc(Left$(h$, 1)) >= 32) Then
                Call OemToChar(h$, h$)
                h$ = h$ + " (" + Mid$(Str$(i%), 2) + ")"
                .AddItem h$
            End If
        End If
    Next i%
End With
    
End Sub

Sub Umspeichern(buf$, pos%)
Dim ind%, ind2%, i%, z%, z1%, z2%
Dim zu!
Dim EK#, VK#
Dim h$, h2$, char$, zuza$, zusa$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

        Ausgabe(pos%).pzn = LiefNr%
        Ausgabe(pos%).Bookmark = FabsRecno&
'        Ausgabe(pos%).key = clsfabs.LastKey(10)
        
        h$ = " "
        FabsErrf% = lZus1.IndexSearch(0, Format(Ausgabe(pos%).pzn, "00000") + "001", FabsRecno&)
        If (FabsErrf% = 0) Then
            Ausgabe(pos%).ZusatzTextKz = 2
            h$ = "+"
        End If
        
        h$ = h$ + vbTab + Lif1.kurz + vbTab + Str$(LiefNr%) + vbTab + Lif1.Name(0) + vbTab + Lif1.Name(1) + vbTab
        h$ = h$ + Lif1.telefon + vbTab + Lif1.Durchwahl
        
        Ausgabe(pos%).Name = h$
End Sub

Function SuchWeiter%(ind%, typ%)
Dim erg%, ind2%
Dim AltRecno&
Dim h$

SuchWeiter% = False

erg% = False

With frmMatchcode.lstSortierung
    AltRecno& = Ausgabe(ind%).Bookmark
    If (typ%) Then
        If (AltRecno& < (.ListCount - 1)) Then
            FabsRecno& = AltRecno& + 1
            erg% = True
        End If
    Else
        If (AltRecno& > 0) Then
            FabsRecno& = AltRecno& - 1
            erg% = True
        End If
    End If
    
    If (erg%) Then
        .ListIndex = FabsRecno&
        h$ = .text
        ind2% = InStr(h$, "(")
        h$ = Mid$(h$, ind2% + 1)
        
        ind2% = Val(Left$(h$, Len(h$) - 1))
        LiefNr% = ind2%
    
        Lif1.GetRecord (LiefNr% + 1)
        SuchWeiter% = True
    End If
End With

End Function

Function SuchArtikel%(Such$, iArbeitAnzzeilen%)
Dim i%, l%, ind%, TextSuche%, IstPzn%, LeftLen%, InhMengeFlag%
Dim InhMenge&
Dim SQLStr$, nSuch$, ch$, h$
Dim iBookmark As Variant

SuchArtikel% = True

OrgSuch$ = Such$

With frmMatchcode.lstSortierung
    FabsRecno& = 0
    For i% = 1 To .ListCount
        .ListIndex = i% - 1
        h$ = .text
        If (h$ >= Such$) Then
            FabsRecno& = i% - 1
            Exit For
        End If
    Next i%
    
    .ListIndex = FabsRecno&
    h$ = .text
    ind% = InStr(h$, "(")
    h$ = Mid$(h$, ind% + 1)
    
    ind% = Val(Left$(h$, Len(h$) - 1))
    LiefNr% = ind%
End With


Lif1.GetRecord (LiefNr% + 1)
Call Umspeichern("", 0)
Call MachAuswahlGrid(iArbeitAnzzeilen%)
Call frmMatchcode.MachMaske
With frmMatchcode.flxarbeit(0)
    .row = 1
    .ColSel = .Cols - 1
    .SetFocus
End With

End Function

Sub MachAuswahlGrid(iArbeitAnzzeilen%)
Dim spBreite%, i%
Dim aFont$

With frmMatchcode.flxarbeit(0)
    .Rows = iArbeitAnzzeilen%
    .FixedRows = 1
            
    .Cols = 7
    .FormatString = "|<Kurz|>Nr|<Name1|<Name2|<Telefon|<Dw"
    frmMatchcode.Font.Bold = True
    .ColWidth(0) = frmMatchcode.TextWidth("X")
    .ColWidth(1) = frmMatchcode.TextWidth("XXXXXXX")
    .ColWidth(2) = frmMatchcode.TextWidth("999")
    .ColWidth(4) = frmMatchcode.TextWidth(String$(25, "X"))
    .ColWidth(5) = frmMatchcode.TextWidth(String$(16, "9"))
    .ColWidth(6) = frmMatchcode.TextWidth(String$(5, "9"))
    frmMatchcode.Font.Bold = False
            
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (i <> 3) Then
            .ColWidth(i%) = .ColWidth(i%) + frmMatchcode.TextWidth("X")
            spBreite% = spBreite% + .ColWidth(i%)
        End If
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(3) = .Width - spBreite%
End With
End Sub


Public Property Get IniSection$()
IniSection$ = INI_SECTION
End Property

Public Property Get InfoSection$()
InfoSection$ = INFO_SECTION
End Property



