Attribute VB_Name = "Module1"
Option Explicit

Sub HolePrivatRezepte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HolePrivatRezepte")
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
End
DefErrEnd:
Dim l&, j&, found&, Max&
Dim NextMulti%, PNr%
Dim h$, Abdatum$, AbZeit$, xc$
Dim IstMulti As Boolean

h$ = Space$(11)
l& = GetPrivateProfileString("Rezeptkontrolle", "Privatrezepte", h$, h$, 11, INI_DATEI)
h$ = Left$(h$, l&)
If Trim(h$) > "" Then
    Abdatum$ = Left$(h$, 6)
    AbZeit$ = Mid$(h$, 7, 4)
Else
    Abdatum$ = "010102"
    AbZeit$ = "0000"
End If


Call VK.GetRecord(1)
xc$ = MKDate(iDate(Abdatum$))
found& = VK.DatumSuche(xc$)
If found& < 0 Then found& = Abs(found&) + 1
Max& = VK.DateiLen / VK.RecordLen

j& = 1
Do While found& <= Max&
    VK.GetRecord (found&)
    IstMulti = False
    If VK.pzn = "9999999" And Mid$(VK.text$, 20, 1) = "x" Then IstMulti = True
    NextMulti% = 1
    If found& < Max& Then
        VK.GetRecord (found& + 1)
        If VK.pzn = "9999999" And Mid$(VK.text$, 20, 1) = "x" Then
            NextMulti% = Val(Mid$(VK.text$, 16, 3))
        End If
        VK.GetRecord (found&)
    End If
    
    If Val(VK.RezEan) = 0 And VK.gebuehren > 0 And (Not IstMulti) Then    'Privatrezept
        RezNr$ = "P"
        ActProgram.VerkPtr = found&
        erg = ActProgram.RezeptHolen(2)
        If erg Then Call ActProgram.WriteRezeptSpeicher
        
    
        PNr% = VK.gebuehren
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + VK.pzn$
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If Not TaxeRec.EOF Then
            If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                PosStr$ = ""
                FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
                If (FabsErrf% = 0) Then
                    ass.GetRecord (FabsRecno& + 1)
                    PosStr$ = Format(ass.poslag, "0")
                End If
                
                AVP# = TaxeRec!VK / 100#
                h$ = Format(TaxeRec!pzn, "0000000") + vbTab + TaxeRec!Name + vbTab + TaxeRec!menge + vbTab + TaxeRec!einheit
                h$ = h$ + vbTab + TaxeRec!HerstellerKB + vbTab + PosStr$
                h$ = h$ + vbTab + Format(AVP#, "0.00") + vbTab + Format(NextMulti%, "0") + vbTab + Format(AVP# * CDbl(NextMulti%), "0.00")
                flxOrg.AddItem h$
            End If
        End If
    End If
    
    found& = found& + 1
    j& = j& + 1
    
    If ((j& Mod 100) = 0) Then
        DoEvents
        If (AbbruchSuche%) Then
            flxOrg.Rows = 1
            Exit Do
        End If
    End If
Loop
Call DefErrPop

End Sub


