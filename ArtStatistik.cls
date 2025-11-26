VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArtStatistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim letztdat%, pp%, xheute%
Dim verbr%(32), datum%(32), qStart%
Dim xAnfang&
Dim Wert!(50)
Dim Jahr!(3)
Dim Wjahr!, WvJahr!
Dim lxheute$

Dim FabsErrf%
Dim FabsRecno&



Private Const DefErrModul = "ARTSTATISTIK.CLS"

Function StatistikRechnen%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StatistikRechnen%")
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
Dim Statistik%, j%, x%, m%, ret%, ok%, dAnzFields%, IstMerkzettelMdb%
Dim iMonat%, iJahr%
Dim xx$, h$, xc$, SQLStr$
Dim AbgabenRec As New ADODB.Recordset

ret% = False

For j% = 0 To 50
    Wert!(j%) = 0
Next j%
For j% = 0 To 3
    Jahr!(j%) = 0
Next j%
Wjahr! = 0
WvJahr! = 0

For j% = 1 To 32
    verbr%(j) = 0: datum%(j) = 0
Next j%

If (Val(pzn$) = 9999999) Then
    StatistikRechnen% = ret%
    Call clsError.DefErrPop
    Exit Function
End If

Statistik% = True
xx$ = ""

h$ = Format(Now, "DDMMYY")
xheute% = clsOpTool.iDate(h$)
lxheute$ = clsOpTool.DateLong(h$)
xAnfang& = Val(Left$(lxheute$, 4)) * 12 + Val(Mid$(lxheute$, 5, 2)) + 1

If (ArtikelDBok) Then
    If (Val(pzn) > 0) Then
        SQLStr = "SELECT * FROM Abgaben WHERE Pzn=" + pzn
        SQLStr = SQLStr + " ORDER BY Datum DESC"
        AbgabenRec.Open SQLStr, ArtikelDB1.ActiveConn
        Do
            If (AbgabenRec.EOF) Then
                Exit Do
            End If
            
            ret = True
            
            iMonat = DateDiff("m", AbgabenRec!datum, Now) + 1
            If (iMonat <= 50) Then
                Wert(iMonat) = Wert(iMonat) + AbgabenRec!AbgabeMenge
            End If
            
            iJahr = Year(Now) - Year(AbgabenRec!datum)
            If (iJahr < 3) Then
                Jahr(iJahr) = Jahr(iJahr) + AbgabenRec!AbgabeMenge
            End If
            
'            If (iJahr = Year(Now)) Then
'                Wjahr = Wjahr + AbgabenRec!AbgabeMenge
'            ElseIf (iJahr = (Year(Now) - 1)) Then
'                WvJahr = WvJahr + AbgabenRec!AbgabeMenge
'            End If
                        
            AbgabenRec.MoveNext
        Loop
        AbgabenRec.Close
    End If
Else
    'pzn$ = "7394479"
    FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Call Ass1.GetRecord(FabsRecno& + 1)
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Call Ast1.GetRecord(FabsRecno& + 1)
        End If
    End If
    
    letztdat% = 0
    qStart% = 0
    
    If (FabsErrf% = 0) Then
        ret% = True
        
    '    x% = Ass1.quarttag: xx$ = clsOpTool.sDate(x%)
    '    If (x% <> 0) Then
    '
    '      '* ersten des Quartals einstellen:
    '      m% = ((Val(Mid$(xx$, 3, 2)) - 1) \ 3) * 3 + 1
    '      xx$ = "01" + Right$("00" + Mid$(Str$(m%), 2), 2) + Right$(xx$, 2)
    '      x% = clsOpTool.iDate(xx$): xc$ = clsOpTool.MKDate(x%)
    '
    '      qStart% = 4
    '      datum%(4) = x%: verbr%(4) = Ass1.quart(3)
    '      datum%(3) = x% - 92: verbr%(3) = Ass1.quart(2)
    '      datum%(2) = x% - 183: verbr%(2) = Ass1.quart(1)
    '      datum%(1) = x% - 275: verbr%(1) = Ass1.quart(0)
    '    End If
    '
    '    For j% = 18 To 1 Step -1
    '      datum%(j% + 4) = Ass1.vkmon(j% - 1)
    '      verbr%(j% + 4) = Ass1.vkmmng(j% - 1)
    '    Next j
    '
    '    letztdat% = 22: Call PosArtikel
    
        x% = Ass1.quarttag: xx$ = clsOpTool.sDate(x%)
        If (x% <> 0) Then
            For j% = 0 To 3
                If (Ass1.quart(j%) = 0) Then
                    x% = 0
                    Exit For
                End If
            Next j%
        End If
            
        If (x% <> 0) Then
            '* ersten des Quartals einstellen:
            m% = ((Val(Mid$(xx$, 3, 2)) - 1) \ 3) * 3 + 1
            xx$ = "01" + Right$("00" + Mid$(Str$(m%), 2), 2) + Right$(xx$, 2)
            x% = clsOpTool.iDate(xx$): xc$ = clsOpTool.MKDate(x%)
            
            qStart% = 4
            datum%(4) = x%: verbr%(4) = Ass1.quart(3)
            datum%(3) = x% - 92: verbr%(3) = Ass1.quart(2)
            datum%(2) = x% - 183: verbr%(2) = Ass1.quart(1)
            datum%(1) = x% - 275: verbr%(1) = Ass1.quart(0)
        End If
        
        For j% = 18 To 1 Step -1
          datum%(j% + qStart%) = Ass1.vkmon(j% - 1)
          verbr%(j% + qStart%) = Ass1.vkmmng(j% - 1)
        Next j
        
        letztdat% = 18 + qStart%: Call PosArtikel
    Else
        On Error Resume Next
        dAnzFields% = MerkzettelDB.TableDefs("Merkzettel").Fields.Count
        IstMerkzettelMdb% = (Err = 0)
        On Error GoTo DefErr
        
        If (IstMerkzettelMdb%) Then
            SQLStr$ = "SELECT * FROM Merkzettel WHERE Pzn='" + clsOpTool.SqlPzn(pzn$, 0) + "'"
            Set MerkzettelRec = MerkzettelDB.OpenRecordset(SQLStr$)
            If (MerkzettelRec.RecordCount > 0) Then
                MerkzettelRec.MoveFirst
                Do
                    If (MerkzettelRec.EOF) Then
                        Exit Do
                    End If
                    
                    m% = MerkzettelRec!lm
                    If (m% <> 0) Then
                        ret% = True
                        
                        pp% = True
                        If (letztdat% >= 32) Then
                            For j% = 2 To 32
                                verbr%(j% - 1) = verbr%(j%)
                                datum%(j% - 1) = datum%(j%)
                            Next j%
                        Else
                            letztdat% = letztdat% + 1
                        End If
                        
                        verbr%(letztdat%) = m%
                        
                        h$ = Format(MerkzettelRec!LieferDatum, "DDMMYY")
                        datum%(letztdat%) = clsOpTool.iDate(h$)
                    End If
                    
                    MerkzettelRec.MoveNext
                Loop
            End If
        Else
            FabsErrf% = Besorgt1.IndexSearch(0, pzn$, FabsRecno&)
            Do
        '        If (FabsErrf%) Or (letztdat% >= 32) Then Exit Do
                If (FabsErrf%) Then Exit Do
        
                Besorgt1.GetRecord (FabsRecno& + 1)
        
                If (pzn$ <> Besorgt1.pzn) Then Exit Do
        
                m% = Besorgt1.lm
                If (m% <> 0) Then
                    ret% = True
                    
                    pp% = True
                    If (letztdat% >= 32) Then
                        For j% = 2 To 32
                            verbr%(j% - 1) = verbr%(j%)
                            datum%(j% - 1) = datum%(j%)
                        Next j%
                    Else
                        letztdat% = letztdat% + 1
                    End If
                    verbr%(letztdat%) = m%
                    datum%(letztdat%) = Besorgt1.dt
                End If
        
                FabsErrf% = Besorgt1.IndexNext(0, FabsRecno&, pzn$, FabsRecno&)
            Loop
        End If
    End If
    
    
    'If (letztdat% = 0) Then Call clsError.DefErrPop: Exit Function
    
    If (ret%) Then
        Call DatSort
        Call MonatsTeilen
    End If
End If

StatistikRechnen% = ret%

Call clsError.DefErrPop
End Function

Sub PosArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PosArtikel")
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
Dim errf%

errf% = 0: pp% = 0
If (InStr(Para1.PosAktivWG, Left$(Ast1.wg, 1)) > 0) Then pp% = True
If (pp%) Then Call VerkaufEinlesen

Call clsError.DefErrPop
End Sub

Sub VerkaufEinlesen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("VerkaufEinlesen")
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
Dim posstart%, j%, x%

posstart% = 0
For j% = 1 To 10
  x% = Ass1.vkmng(j% - 1)
  If (x% <> 0) Then
    letztdat% = letztdat% + 1: posstart% = True
    verbr%(letztdat%) = x%
    datum%(letztdat%) = Ass1.vkdat(j% - 1)
  End If
Next j%

Call clsError.DefErrPop
End Sub

Sub DatSort()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DatSort")
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
Dim j%, k%, ende%, temp%

For j% = 1 To (letztdat% - 1)
  For k% = (j% + 1) To letztdat%
    If (datum%(j%) = 0) Or (verbr%(j%) = 0) Or (datum%(k%) <> 0 And datum%(k%) < datum%(j%)) Then
      temp% = datum%(j%)
      datum%(j%) = datum%(k%)
      datum%(k%) = temp%
      temp% = verbr%(j%)
      verbr%(j%) = verbr%(k%)
      verbr%(k%) = temp%
    End If
  Next k%
Next j%
j% = 1: ende% = 0
While (j% <= letztdat%) And Not ende%
  If (verbr%(j%) = 0) Or (datum%(j%) = 0) Then ende% = True: letztdat% = j% - 1
  j% = j% + 1
Wend

Call clsError.DefErrPop
End Sub

Sub MonatsTeilen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MonatsTeilen")
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
Dim j%, tt1%, mm1%, WertMonat%, altdatum%, x%, tt%, mm%, zeit%, Erster%, Letzter%, Monat%, MaxMonat%
Dim jj1&, jj&, Heuer&, Termin&, Jahr&, MaxJahr&
Dim wwert!, TagVerbr!, max!
Dim xc$, xx$, h$


tt1% = Val(Right$(lxheute$, 2))
mm1% = Val(Mid$(lxheute$, 5, 2))
jj1& = Val(Left$(lxheute$, 4))
wwert! = 0!
WertMonat% = 1
altdatum% = xheute%
xAnfang& = jj1& * 12 + mm1% + 1
j% = letztdat%

While (j% >= 1)
  x% = datum%(j%):  xc$ = clsOpTool.MKDate(x%): xx$ = clsOpTool.sDate(x%): xx$ = clsOpTool.DateLong(xx$)
  tt% = Val(Right$(xx$, 2)): mm% = Val(Mid$(xx$, 5, 2)): jj& = Val(Left$(xx$, 4))
  If (pp% <> 0) And (j% > qStart%) Then
    jj1& = jj&: mm1% = mm%
    WertMonat% = xAnfang& - (jj1& * 12 + mm1%)
    If (WertMonat% < 0) Then WertMonat% = 0
    If (WertMonat% <= 50) Then
      Wert!(WertMonat%) = Wert!(WertMonat%) + verbr%(j%)
    End If
  ElseIf (mm% = mm1%) And (jj& = jj1&) Then
    wwert! = wwert! + verbr%(j%)
  Else
    zeit% = altdatum% - datum%(j%)
    TagVerbr! = CSng(verbr%(j%)) / CSng(zeit% + tt1%)
    WertMonat% = xAnfang& - (jj1& * 12 + mm1%)
    If (WertMonat% < 0) Then WertMonat% = 0
    Wert!(WertMonat%) = Wert!(WertMonat%) + wwert! + TagVerbr! * CSng(tt1%)
    Letzter% = altdatum% - tt1% + 1
    mm1% = mm1% - 1: If (mm1% < 1) Then jj1& = jj1& - 1: mm1% = 12
    While (mm1% <> mm% Or jj1& <> jj&) And WertMonat% <= 25
      WertMonat% = xAnfang& - (jj1& * 12 + mm1%)
      If (WertMonat% < 0) Then WertMonat% = 0
      xx$ = "01" + Right$(Str$(mm1%), 2) + Right$(Str$(jj1&), 2)
      x% = 1
      While (x% <> 0)
        x% = InStr(xx$, " "): If (x% > 0) Then Mid$(xx$, x%, 1) = "0"
      Wend
      x% = clsOpTool.iDate(xx$): Erster% = x%
      Wert!(WertMonat%) = Wert!(WertMonat%) + TagVerbr! * CSng(Letzter% - Erster%)
      Letzter% = Erster%
      mm1% = mm1% - 1: If (mm1% < 1) Then jj1& = jj1& - 1: mm1% = 12
    Wend
    wwert! = TagVerbr! * CSng(Letzter% - datum%(j%))
  End If
  altdatum% = datum%(j%): tt1% = tt%: mm1% = mm%: jj1& = jj&
  j% = j% - 1
  If (WertMonat% > 25) Then j% = 0
Wend
If (pp% = 0) Then
  WertMonat% = xAnfang& - (jj1& * 12 + mm1%)
  If (WertMonat% < 0) Then WertMonat% = 0
  If (WertMonat% >= 0) And (WertMonat% <= 25) Then Wert!(WertMonat%) = wwert!
End If

For j% = 0 To 50
    Wert!(j%) = Int(Wert!(j%) + 0.501)
Next j%

max! = 0: Wjahr! = 0: WvJahr! = 0
Heuer& = Val(Left$(lxheute$, 4))
For j% = 1 To 25
  Termin& = xAnfang& - j%: Jahr& = (Termin& - 1) \ 12: Monat% = Termin& - Jahr& * 12
  If (Jahr& = Heuer&) Then Wjahr! = Wjahr! + Wert!(j%)
  If (Jahr& = Heuer& - 1) Then WvJahr! = WvJahr! + Wert!(j%)
  If (Wert!(j%) > max!) Then max! = Wert!(j%): MaxJahr& = Jahr&: MaxMonat% = Monat%
Next j%

Call clsError.DefErrPop
End Sub

Public Property Get MonatsWert!(ByVal ind%)
If (ind% < 0) Then
    MonatsWert! = 0
Else
    MonatsWert! = Wert!(ind%)
End If
End Property

Public Property Get JahresWert!(ByVal ind%)
Dim iJahr%
Dim Heuer&

If (ArtikelDBok) Then
    iJahr = Year(Now) - ind
    If (iJahr < 3) Then
        JahresWert = Jahr(iJahr)
    End If
Else
    Heuer& = Val(Left$(lxheute$, 4))
    If (ind% = Heuer&) Then
        JahresWert! = Wjahr!
    ElseIf (ind% = (Heuer& - 1)) Then
        JahresWert! = WvJahr!
    Else
        JahresWert! = 0!
    End If
End If

End Property

Public Property Get Anfang&()
Anfang& = xAnfang&
End Property

