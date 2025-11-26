Attribute VB_Name = "modBesorger"
Option Explicit

Type VerkaufStruct
  RezNr       As Integer
  EndeT       As Byte
  pzn         As String * 7
  text        As String * 36
  preis       As String * 9
  mw          As String * 1
  bs          As Byte
  Datum       As String * 2
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


Public VkRec As VerkaufStruct

Public BesorgerAconto$
Public BesorgerInfo$(9)
Public BesorgerMenge%

Dim ANFERT%
Dim AnfBlock%

Private Const DefErrModul = "wabholer.bas"

'Function AbholerNummer%(Nummer%, pzn$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("AbholerNummer%")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, satz1%, satz%, header%, bit%, bitmaske%, frei%, kistefertig%, allesunfertig%, Menge%
'Dim text%, InfoText%, PznGefunden%
'Dim KistenSatz&, BlockNummer&
'Dim preis#
'Dim ch$, Arttext$, aiPzn$, aconto$, txt$, wg$, mText$, mx$, h$, h1$, h2$
'
'AbholerNummer% = False
'
'If Abs(Nummer%) > 999 Then
'    'Call FehlerKastl(seite%, " Abholnummer" + str$(Nummer%) + " ist zu groﬂ!")
'    'Nummer% = 0
'    Call DefErrPop: Exit Function
'End If
'
'If Nummer% < 0 Then
'    Call DefErrPop: Exit Function
'End If
'
'clsAnfert1.OpenDatei
'clsAnfBlock1.OpenDatei
'ANFERT% = FileOpen("ANFERT.DAT", "R")
'AnfBlock% = FileOpen("ANFBLOCK.DAT", "R")
''If g.MehrPlatz% Then Lock #f.af%, 1 To Len(af)
'If LOF(ANFERT%) = 0 Then
'    Call DefErrPop: Exit Function
'End If
'
'satz% = ((Nummer% - 1) \ 8) + Len(af) + 1
'ch$ = Chr$(0)
'Get #ANFERT%, satz%, ch$
'
'satz1% = (Nummer% - 1) \ 8
'satz% = satz1% \ 48
'clsAnfert1.GetRecord (satz% + 2)
'h$ = clsAnfert1.wert("ERST")
'ch$ = Mid$(h$, (satz1% Mod 48) + 1, 1)
'
'header% = Asc(ch$)
'bit% = 7 - ((Nummer% - 1) Mod 8)
'bitmaske% = 2 ^ bit%
'frei% = ((header% And bitmaske%) = 0)
'
'If Not (frei%) Then
'    '* Kisteninfo
''    KistenSatz& = Nummer% + 6
''    Get #ANFERT%, Len(af) * (KistenSatz&) + 1, af
'
'    clsAnfert1.GetRecord (Nummer% + 7)
'
'    kistefertig% = 3
'    allesunfertig% = True
'
'    PznGefunden% = False
'    For i% = 0 To 9
''        BlockNummer& = af.AnfBlock%(i%)
'        BlockNummer& = clsAnfert1.wert("ANFBLOCK", i%)
'        If BlockNummer& > 0 Then
'            afVonWann$ = clsAnfert1.wert("VONWANN")
''            Get #AnfBlock%, Len(aw) * ((BlockNummer& - 1) * 11 + 40) + 1, aw
'            clsAnfBlock1.GetRecord (((BlockNummer& - 1) * 11 + 40) + 1)
'            If InStr("RTHBP", clsAnfBlock1.wert("WASTUN")) > 0 And Asc(clsAnfBlock1.wert("STATUS")) > 0 Then
'                For j% = 1 To 5
''                    Get #AnfBlock%, Len(aw) * ((BlockNummer& - 1) * 11 + 40 + j%) + 1, ai(0)
'                    clsAnfBlock1.GetRecord (((BlockNummer& - 1) * 11 + 40 + j%) + 1)
''                    h$ = RTrim(ai(0).InfoZeile)
'                    h$ = RTrim(clsAnfBlock1.wert("INFOZEILE"))
'                    If Mid$(h$, 14, 4) = "PZN=" Then
'                        aiPzn$ = Mid$(h$, 18, 7)
'                        If (aiPzn$ = pzn$) Then
'                            PznGefunden% = True
'                            Exit For
'                        End If
'                    End If
'                Next j%
'            End If
'            If (PznGefunden%) Then Exit For
'        End If
'    Next i%
'
'    If (PznGefunden%) Then
'        For j% = 1 To 10
''            Get #AnfBlock%, Len(aw) * ((BlockNummer& - 1) * 11 + 40 + j%) + 1, ai(j%)
'            clsAnfBlock1.GetRecord (((BlockNummer& - 1) * 11 + 40 + j%) + 1)
'            h$ = clsAnfBlock1.wert("INFOZEILE")
''            h$ = ai(j%).InfoZeile
'            h1$ = Mid$(h$, 6, 2)
'            h2$ = Left$(h$, 7)
'            If (j% <= 5) And (h1$ = "A=" Or h1$ = "G=" Or h2$ = "ACONTO=" Or h2$ = "GEB‹HR=") Then
''                wg$ = Mid$(ai.InfoZeile$, 2, 2)
'                BesorgerAconto$ = Mid$(h$, 8, 10)
'                BesorgerMenge% = Val(Mid$(h$, 21, 4))
'                If (BesorgerMenge% = 0) Then BesorgerMenge% = 1
'            End If
'        Next j%
'        AbholerNummer% = True
'    End If
'
'End If
''If g.MehrPlatz% Then Unlock #f.af%, 1 To Len(af)
'
'Close #ANFERT%
'Close #AnfBlock%
'
'Call DefErrPop
'End Function

Function AbholerNummer%(Nummer%, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerNummer%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, satz1%, satz%, header%, bit%, bitmaske%, frei%, kistefertig%, allesunfertig%, Menge%
Dim text%, InfoText%, PznGefunden%, belegt%, BlockNummer%
Dim KistenSatz&
Dim preis#
Dim ch$, Arttext$, aiPzn$, aconto$, txt$, wg$, mText$, mx$, h$, h1$, h2$

AbholerNummer% = False

If Abs(Nummer%) > 999 Then
    'Call FehlerKastl(seite%, " Abholnummer" + str$(Nummer%) + " ist zu groﬂ!")
    'Nummer% = 0
    Call DefErrPop: Exit Function
End If

If Nummer% < 0 Then
    Call DefErrPop: Exit Function
End If

kiste.OpenDatei
belegt% = kiste.belegt(Nummer%)
If (belegt%) Then
    BlockNummer% = kiste.PznInKiste(Nummer%, pzn$)
    If (BlockNummer% >= 0) Then
        kiste.GetInhalt (BlockNummer%)
        For j% = 0 To 9
            h$ = RTrim$(kiste.InfoText(j%))
            BesorgerInfo$(j%) = h$
            h1$ = Mid$(h$, 6, 2)
            h2$ = Left$(h$, 7)
            If (j% <= 5) And (h1$ = "A=" Or h1$ = "G=" Or h2$ = "ACONTO=" Or h2$ = "GEB‹HR=") Then
'                wg$ = Mid$(ai.InfoZeile$, 2, 2)
                BesorgerAconto$ = Mid$(h$, 8, 10)
                BesorgerMenge% = Val(Mid$(h$, 21, 4))
                If (BesorgerMenge% = 0) Then BesorgerMenge% = 1
            End If
        Next j%
        AbholerNummer% = True
    End If
End If

kiste.CloseDatei

Call DefErrPop
End Function

Sub SucheBesorgerInVk(AbholNr%, PersCode%, KundenNr%)
Dim i%, j%, VERKAUF%
Dim von&, bis&, Index&, VerkaufMax&, ret&
Dim h$, s$, SuchDatum$, SuchZeit$, Such$, SuchKiste$

PersCode% = 0
KundenNr% = 0

SuchDatum$ = Left$(kiste.VonWann, 2)
SuchZeit$ = MKI(Asc(Mid$(kiste.VonWann, 3, 1)) * 100 + Asc(Mid$(kiste.VonWann, 4, 1)))
Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)

SuchKiste$ = Format(AbholNr%, "0")

For i% = 1 To 4
    Debug.Print Asc(Mid$(Such$, i%, 1));
Next i%
Debug.Print
Debug.Print
                
h$ = "verkauf.dat"
VERKAUF% = FileOpen(h$, "R", "R", 128)

VerkaufMax& = (LOF(VERKAUF%) / 128) - 1

For j% = 0 To 1
    If (j% = 0) Then
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Mid$(SuchZeit$, 1, 1)
    Else
        Such$ = SuchDatum$ + Mid$(SuchZeit$, 2, 1) + Chr$(Asc(Mid$(SuchZeit$, 1, 1)) - 1)
    End If
    
    von& = 1&
    bis& = VerkaufMax&
    
    Do While (von& <= bis&)
        Index& = (von& + bis&) \ 2
        Get #VERKAUF%, Index& + 1, VkRec
        s$ = VkRec.Datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        Debug.Print Index&;
        For i% = 1 To 4
            Debug.Print Asc(Mid$(s$, i%, 1));
        Next i%
        Debug.Print
        If (Such$ = s$) Then
            ret& = Index&
            Exit Do
        ElseIf (Such$ < s$) Then
            bis& = Index& - 1
        Else
            von& = Index& + 1
        End If
    Loop
    
    If (ret& >= 1) Then Exit For
Next j%

If (ret& >= 1) Then
    Do
        ret& = ret& - 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.Datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        If (Such$ <> s$) Then
            Exit Do
        End If
    Loop
    
    Do
        ret& = ret& + 1
        Get #VERKAUF%, ret& + 1, VkRec
        s$ = VkRec.Datum + Mid$(VkRec.zeit, 2, 1) + Mid$(VkRec.zeit, 1, 1)
        If (Such$ <> s$) Then
            Exit Do
        End If
        If (InStr(VkRec.text, "Abhol-Nr") > 0) And (InStr(VkRec.text, SuchKiste$) > 0) Then
            KundenNr% = VkRec.knr
            PersCode% = VkRec.PersCode
            Exit Do
        End If
    Loop
End If

Close #VERKAUF%

End Sub
    

