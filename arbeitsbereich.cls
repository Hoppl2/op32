VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArbeitBereich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ParentFlex As MSFlexGrid

Dim iAnzArbeitSpalten%
Dim iDatei$(10)
Dim iKurz$(10)
Dim iBezeichnung$(10)
Dim iLaenge%(10)
Dim iAusrichtung(10) As Byte
Dim iKzBelegung$(10)

Dim iIniDatei$
Dim iIniSection$

Dim KlassenCheck%

Const MAX_ARBEIT_SPALTEN = 10


Private Const DefErrModul = "ARBEITSBEREICH.CLS"

Sub InitArbeitBereich(iFlex As Object, IniDatei$, IniSection$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitArbeitBereich")
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
Dim i%, j%, ind%
Dim l&
Dim h$, h2$, key$

Set ParentFlex = iFlex
iIniDatei$ = IniDatei$
iIniSection$ = IniSection$


For i% = 1 To MAX_ARBEIT_SPALTEN
    iDatei$(i% - 1) = ""
    iKurz$(i% - 1) = ""
    iBezeichnung$(i% - 1) = ""
        
    h$ = Space$(100)
    key$ = "Spalte" + Format(i%, "0")
    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
    h$ = Trim$(Left$(h$, l&))
    If (Len(h$) > 0) Then
        iAnzArbeitSpalten% = i%
        ind% = InStr(h$, ",")
        h2$ = ""
        If (ind% > 0) Then
            h2$ = Mid$(h$, ind% + 1)
            h$ = Left$(h$, ind% - 1)
        End If
        ind% = InStr(h$, ".")
        If (ind% > 0) Then
            iDatei$(i% - 1) = UCase$(Left$(h$, ind% - 1))
            iKurz$(i% - 1) = Mid$(h$, ind% + 1)
            
            ind% = InStr(h2$, ",")
            If (ind% > 0) Then
                h$ = Left$(h2$, ind% - 1)
                h2$ = Mid$(h2$, ind% + 1)
            Else
                h$ = h2$
                h2$ = "1,0"
            End If
'            iBezeichnung$(i% - 1) = clsDat.firstlettersucase(h$)
            iBezeichnung$(i% - 1) = h$
            
            ind% = InStr(h2$, ",")
            If (ind% > 0) Then
                h$ = Left$(h2$, ind% - 1)
                h2$ = Mid$(h2$, ind% + 1)
            Else
                h$ = h2$
                h2$ = "0"
            End If
            iLaenge%(i% - 1) = Val(h$)
            
            
            ind% = InStr(h2$, ",")
            If (ind% > 0) Then
                h$ = Left$(h2$, ind% - 1)
                h2$ = Mid$(h2$, ind% + 1)
            Else
                h$ = h2$
                h2$ = ""
            End If
            iAusrichtung(i% - 1) = Val(h$)
            iKzBelegung$(i% - 1) = h2$
        End If
    End If
Next i%

KlassenCheck = 0

Call clsError.DefErrPop
End Sub

Function ZeigeArbeitBereich$(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeArbeitBereich$")
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
Dim i%, j%, ind%
Dim h$, txWert$, tx$
Dim ret$

ret$ = ""

If (KlassenCheck = 0) Then
'    CheckArbeitBereich
    For j% = 1 To iAnzArbeitSpalten%
        h$ = iDatei$(j% - 1)
        If (h$ <> "") Then
            If (h$ = "KZ") Then
'                If (clsKz.CheckKzKlasse(h$) = 0) Then
'                    iDatei$(j% - 1) = ""
'                    iKurz$(j% - 1) = ""
'                End If
            Else
                If (clsDat.CheckKlasse(h$) = 0) Then
                    iDatei$(j% - 1) = ""
                    iKurz$(j% - 1) = ""
                End If
            End If
        End If
    Next j%
    KlassenCheck = True
End If
            
ArtNeu% = False
AstNeu% = False
AssNeu% = False
TaxeNeu% = False
LiefNeu% = False
hTaxNeu% = False
'If (AuchWerte% = False) Then
'    AstNeu% = True
'    AssNeu% = True
'    TaxeNeu% = True
'    LiefNeu% = True
'End If

ArtDa% = False
AstDa% = False
AssDa% = False
TaxeDa% = False
LiefDa% = False
hTaxDa% = False
SonderPrDa% = 0


For j% = 1 To iAnzArbeitSpalten%
    h$ = iDatei$(j% - 1)
    If (h$ <> "") Then
        If (h$ = "KZ") Then
            Call clsKz.CheckKzDS(pzn$, iKzBelegung$(j% - 1))
        Else
            Call clsDat.CheckDateiDS(pzn$, h$, False)
        End If
    End If
Next j%

For j% = 1 To iAnzArbeitSpalten%
    ret$ = ret$ + vbTab
    
    h$ = iDatei$(j% - 1)
    If (h$ = "KZ") Then
        txWert$ = clsKz.HoleKzValue(iKzBelegung$(j% - 1), pzn$)
    Else
        txWert$ = clsDat.HoleDateiValue(h$, iKurz$(j% - 1))
    End If
    ret$ = ret$ + txWert$
Next j%

ZeigeArbeitBereich$ = ret$

Call clsError.DefErrPop
End Function

Sub LoescheArbeitBelegung(col%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoescheArbeitBelegung")
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
Dim i%, j%, k%, ZeileLeer%
Dim h$

If (iAnzArbeitSpalten% > 0) Then
    For j% = col% To (iAnzArbeitSpalten% - 2)
        iDatei$(j%) = iDatei$(j% + 1)
        iKurz$(j%) = iKurz$(j% + 1)
        iBezeichnung$(j%) = iBezeichnung$(j% + 1)
        iLaenge%(j%) = iLaenge%(j% + 1)
        iAusrichtung(j%) = iAusrichtung(j% + 1)
        iKzBelegung$(j%) = iKzBelegung$(j% + 1)
    Next j%
    iDatei$(MAX_ARBEIT_SPALTEN - 1) = ""
    iKurz$(MAX_ARBEIT_SPALTEN - 1) = ""
    iBezeichnung$(MAX_ARBEIT_SPALTEN - 1) = ""
    iLaenge%(MAX_ARBEIT_SPALTEN - 1) = 0
    iAusrichtung(MAX_ARBEIT_SPALTEN - 1) = 0
    iKzBelegung$(MAX_ARBEIT_SPALTEN - 1) = ""
    iAnzArbeitSpalten% = iAnzArbeitSpalten% - 1
End If
  
Call SpeicherArbeitBelegung

Call clsError.DefErrPop
End Sub

Sub InsertArbeitBelegung(col%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InsertArbeitBelegung")
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
Dim j%, k%

If (iAnzArbeitSpalten% < MAX_ARBEIT_SPALTEN) Then
    For j% = (iAnzArbeitSpalten% - 1) To col% Step -1
        iDatei$(j% + 1) = iDatei$(j%)
        iKurz$(j% + 1) = iKurz$(j%)
        iBezeichnung$(j% + 1) = iBezeichnung$(j%)
        iLaenge%(j% + 1) = iLaenge%(j%)
        iAusrichtung(j% + 1) = iAusrichtung(j%)
        iKzBelegung$(j% + 1) = iKzBelegung$(j%)
    Next j%
    iDatei$(col%) = ""
    iKurz$(col%) = ""
    iBezeichnung$(col%) = ""
    iLaenge%(col%) = 0
    iAusrichtung(col%) = 0
    iKzBelegung$(col%) = ""
    iAnzArbeitSpalten% = iAnzArbeitSpalten% + 1
    
    Call SpeicherArbeitBelegung
End If
        
Call clsError.DefErrPop
End Sub

Function EditArbeitBelegung%(col%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditArbeitBelegung%")
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
Dim row%, ind%, ret%
Dim OrgInfoLayoutBelegung$, h$, h2$, s$

ret% = False

'col% = (ParentFlex.col - 4) \ 2
InfoLayoutDatei$ = iDatei$(col%)
If (iDatei$(col%) = "KZ") Then
    InfoLayoutBelegung$ = ""
    h$ = iKzBelegung$(col%)
    If (Right$(h$, 1) <> ",") Then h$ = h$ + ","
    Do
        If (h$ = "") Then Exit Do
        ind% = InStr(h$, ",")
        s$ = Left$(h$, ind% - 1)
        h$ = Mid$(h$, ind% + 1)
    
        s$ = iDatei$(col%) + "." + UCase$(s$)
        InfoLayoutBelegung$ = InfoLayoutBelegung$ + s$ + " "
    Loop
Else
    InfoLayoutBelegung$ = iDatei$(col%) + "." + iKurz$(col%)
End If
OrgInfoLayoutBelegung$ = InfoLayoutBelegung$
Load frmInfoLayout
With ParentFlex.Container
    frmInfoLayout.Left = .Left + (.Width - frmInfoLayout.Width) / 2
    frmInfoLayout.Top = .Top + (.Height - frmInfoLayout.Height) / 2
End With
frmInfoLayout.Show 1
If (OrgInfoLayoutBelegung$ <> InfoLayoutBelegung$) Then
    ret% = True
    ind% = InStr(InfoLayoutBelegung$, ".")
    iDatei$(col%) = Left$(InfoLayoutBelegung$, ind% - 1)
    If (iDatei$(col%) = "KZ") Then
        iKurz$(col%) = "kz"
        InfoLayoutBelegung$ = InfoLayoutBelegung$ + " "
        h$ = ""
        Do
            If (Trim(InfoLayoutBelegung$) = "") Then Exit Do
            ind% = InStr(InfoLayoutBelegung$, " ")
            s$ = Left$(InfoLayoutBelegung$, ind% - 1)
            InfoLayoutBelegung$ = Mid$(InfoLayoutBelegung$, ind% + 1)
            ind% = InStr(s$, ".")
            h$ = h$ + Mid$(s$, ind% + 1) + ","
        Loop
        iKurz$(col%) = "kz"
        iKzBelegung$(col%) = h$
    Else
        iKurz$(col%) = Mid$(InfoLayoutBelegung$, ind% + 1)
        iKzBelegung$(col%) = ""
    End If
    
    
    iBezeichnung$(col%) = InfoLayoutBezeichnung$ '""
    iLaenge%(col%) = 10
    iAusrichtung(col%) = 0
    
    Call SpeicherSpaltenBelegung(col%)
End If

EditArbeitBelegung% = ret%

Call clsError.DefErrPop
End Function

Sub SpeicherArbeitBelegung()
Dim i%, j%
Dim l&
Dim h$, key$

For i% = 1 To MAX_ARBEIT_SPALTEN
    Call SpeicherSpaltenBelegung(i% - 1)
Next i%

End Sub

Private Sub SpeicherSpaltenBelegung(col%)
Dim l&
Dim h$, key$

If (col% < iAnzArbeitSpalten%) Then
    h$ = iDatei$(col%) + "." + iKurz$(col%) + ","
    h$ = h$ + iBezeichnung$(col%) + "," + Trim(Str$(iLaenge%(col%))) + ","
    h$ = h$ + Trim(Str$(iAusrichtung(col%)))
    If (iKzBelegung$(col%) <> "") Then
        h$ = h$ + "," + iKzBelegung$(col%)
    End If
Else
    h$ = ""
End If

key$ = "Spalte" + Format(col% + 1, "0")
l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)

End Sub

Public Property Get Bezeichnung$(ByVal col%)
Bezeichnung$ = iBezeichnung$(col%)
End Property
Public Property Let Bezeichnung(ByVal col%, ByVal NeuerWert$)
iBezeichnung$(col%) = NeuerWert$
Call SpeicherSpaltenBelegung(col%)
End Property

Public Property Get Laenge%(ByVal col%)
Laenge% = iLaenge%(col%)
End Property
Public Property Let Laenge(ByVal col%, ByVal NeuerWert%)
iLaenge%(col%) = NeuerWert%
Call SpeicherSpaltenBelegung(col%)
End Property

Public Property Get Ausrichtung$(ByVal col%)
Ausrichtung$ = iAusrichtung(col%)
End Property
Public Property Let Ausrichtung(ByVal col%, ByVal NeuerWert$)
iAusrichtung(col%) = NeuerWert$
Call SpeicherSpaltenBelegung(col%)
End Property

Public Property Get IstKennzeichen%(ByVal col%)
Dim ret%
ret% = False
If (iDatei$(col%) = "KZ") Then ret% = True
If (InStr(iKzBelegung$(col%), "SONDERPR") > 0) Then ret% = 0
IstKennzeichen% = ret%
End Property

Public Property Get AnzArbeitSpalten%()
AnzArbeitSpalten% = iAnzArbeitSpalten%
End Property


