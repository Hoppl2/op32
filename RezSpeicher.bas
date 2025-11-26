Attribute VB_Name = "modRezSpeicher"
Option Explicit

Public Const REZ_SPEICHER = "rezepte.mdb"
Public Const MARS_REZ_SPEICHER = "Mars_rezepte.mdb"
Const KASSEN_DB = "Kassen.mdb"

Public Const MAX_KKTYP = 100 ' 25

Public Kkasse As clsKkassen
Public KKSatz$
Public AbrechMo As String

Public RezSpeicherDB As Database
Public KassenDB As Database

Public RezepteTd As TableDef
Public ArtikelTd As TableDef
Public AuswertungTd As TableDef
Public AbrechDatenTd As TableDef
Public AutIdemKkTd As TableDef

Public KasseTd As TableDef

Public RezepteRec As Recordset
Public ArtikelRec As Recordset
Public AuswertungRec As Recordset
Public KasseRec As Recordset
Public AbrechDatenRec As Recordset
Public AutIdemKkRec As Recordset
Public ParenteralTaxierungRec As Recordset
Public ParenteralTmRec As Recordset
Public AbrechnungsMeldungenRec As Recordset

Public RezSpeicherOK%

Public RezHistorieKassenNr$
Public RezHistorieKassenName$
Public RezHistorieDatum$
Public RezHistorieIndexSuche%
Public RezHistorieTagDirekt%
Public RezHistorieTaxSumme#

Public DarfRezSpeicher As Boolean
Public DarfImportKontrolle As Boolean
Public AbrechMonat$
Public AbrechAbgabeDatum$

Public RezSpeicherModus%

Public ImpAlternativModus%
Public ImpAlternativPara$(2)

Public ArtIndexDebug%
Public TmCheck%
Public BenutzerSignatur%

Public Const MARS_REZEPT_KONTROLLE = 1
Public Const MARS_REZEPT_DRUCK = 2


Public MarsRezSpeicherDB As Database
Public MarsRezepteRec As Recordset
Public MarsArtikelRec As Recordset
Public MarsAnmerkungenRec As Recordset

Public MarsModus%
Public MarsAutomaticModus%
Public MarsAutomaticDatum(1) As Date
Public MarsRezeptAnmerkungen$
Public MarsRezeptZurückgestellt As Boolean

Public SubstitutionsMg As Double
Public SubstitutionsAbgaben(2) As Integer
Public SubstitutionsEinzelpreis As Double


Type RectType
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const DefErrModul = "REZSPEICHER.BAS"

Sub AbrechMonatErmitteln(Optional Datum As String = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbrechMonatErmitteln")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, ok%, mm%, iTyp%
Dim dat1$, dat2$, dat3$, jj$, h$
Dim wann As Date

If Datum = "" Then
  wann = Now
Else
  wann = DateValue(Datum)
End If
ok% = True
AbrechMonat$ = Format(wann, "YYMM")
AbrechAbgabeDatum$ = AbrechMonat$ + "01"
i% = Format(wann, "mm") - 1
If i% = 0 Then i% = 12
With AbrechDatenRec
    .MoveFirst
    Do While Not .EOF
        If AbrechDatenRec!Unique = i% Then
            dat1$ = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
            jj$ = Format(wann, "YY")
            If (i% = 12) Then
                jj$ = Format(Val(jj$) - 1, "00")
            End If
            dat2$ = "01" + "." + Format(AbrechDatenRec!Unique, "00") + "." + jj$
            
            mm% = AbrechDatenRec!Unique + 2
            If (mm% > 12) Then
                mm% = mm% - 12
                jj$ = Format(Val(jj$) + 1, "00")
            End If
            dat3$ = "01" + "." + Format(mm%, "00") + "." + jj$
            
            If DateValue(CDate(dat1$)) < DateValue(CDate(dat2$)) Then
                Call MessageBox("Achtung: Bitte unbedingt Abrechnungsdatum für aktuelles Jahr eintragen !" + vbCrLf + "(Extras -> Optionen -> Abrechnungsdaten)", vbCritical)
                ok% = False
            ElseIf DateValue(CDate(dat1$)) >= DateValue(CDate(dat3$)) Then
                Call MessageBox("Achtung: Bitte unbedingt Abrechnungsdatum für " + Mid$(dat2$, 4) + " eintragen !" + vbCrLf + "(Extras -> Optionen -> Abrechnungsdaten)", vbCritical)
                ok% = False
            ElseIf DateValue(CDate(dat1$)) < DateValue(wann) Then
                If (AbrechDatenRec!Unique = 12) Then
                    .MoveFirst
                Else
                    .MoveNext
                End If
            
                If Not (.EOF) Then
                    dat1$ = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
                    jj$ = Format(wann, "YY")
                    dat2$ = "01" + "." + Format(AbrechDatenRec!Unique, "00") + "." + jj$
                End If
                
                If (.EOF) Or (DateValue(CDate(dat1$)) < DateValue(CDate(dat2$))) Then
                    Call MessageBox("Achtung: Bitte unbedingt Abgabedaten für aktuelles Jahr eintragen !" + vbCrLf + "(Extras -> Optionen -> Abrechnungsdaten)", vbCritical)
                    ok% = False
                ElseIf DateValue(CDate(dat1$)) < DateValue(wann) Then
                    If (AbrechDatenRec!Unique = 12) Then
                        .MoveFirst
                    Else
                        .MoveNext
                    End If
                
                    If Not (.EOF) Then
                        dat1$ = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
                        jj$ = Format(wann, "YY")
                        dat2$ = "01" + "." + Format(AbrechDatenRec!Unique, "00") + "." + jj$
                    End If
                    
                    If (.EOF) Or (DateValue(CDate(dat1$)) < DateValue(CDate(dat2$))) Then
                        Call MessageBox("Achtung: Bitte unbedingt Abrechnungsdatum für aktuelles Jahr eintragen !" + vbCrLf + "(Extras -> Optionen -> Abrechnungsdaten)", vbCritical)
                        ok% = False
                    End If
                End If
            End If


'            Do While DateValue(CDate(Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2))) < DateValue(wann) And Not .EOF
'                .MoveNext
'            Loop

            If (ok%) And (Not .EOF) Then
                AbrechAbgabeDatum$ = Mid(AbrechDatenRec!Datum, 5, 2) + Mid(AbrechDatenRec!Datum, 3, 2) + Left(AbrechDatenRec!Datum, 2)
                AbrechMonat$ = Format(AbrechDatenRec!Unique, "00")
                j% = Val(Format(wann, "YY"))
                If AbrechDatenRec!Unique = 12 And Month(wann) < 11 Then
                    j% = j% - 1
                End If
                AbrechMonat$ = Format(j%, "00") + AbrechMonat$
            End If
            Exit Do
        End If
        .MoveNext
    Loop
End With
h$ = "Rezeptkontrolle - Abrechnungsmonat "
If (ok%) Then
    h$ = h$ + Format(CDate("01." + Mid(AbrechMonat$, 3, 2) + "." + Left(AbrechMonat$, 2)), "mmmm") + " bis " + Mid(AbrechAbgabeDatum$, 5, 2) + "." + Mid(AbrechAbgabeDatum$, 3, 2) + "." + Left(AbrechAbgabeDatum$, 2)
Else
    h$ = h$ + "UNBEKANNT !"
End If
frmAction.Caption = h$
'frmAction.Caption = "Rezeptkontrolle - Abrechnungsmonat " + Format(CDate("01." + Mid(AbrechMonat$, 3, 2) + "." + Left(AbrechMonat$, 2)), "mmmm") + " bis " + Mid(AbrechAbgabeDatum$, 5, 2) + "." + Mid(AbrechAbgabeDatum$, 3, 2) + "." + Left(AbrechAbgabeDatum$, 2)
Call DefErrPop
End Sub

Sub AddImpErspart()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AddImpErspart")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Unique As String
Dim h$, SQLStr$
Dim AnzRezeptArtikel%, j%
Dim aktvk As Double, originalvk As Double
Dim pzn As String, DruckDat As String

        Call WinArtDebug("100")

Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)

RezepteRec.index = "Unique"
RezepteRec.Seek ">=", "0407"
Do While Not RezepteRec.EOF
  If Left(sCheckNull(RezepteRec!Unique), 4) >= "0407" And Val(Left(sCheckNull(RezepteRec!Unique), 4)) > 0 Then
    
    RezepteRec.Edit
    RezepteRec!ImpErspart = 0
        
    AnzRezeptArtikel% = dCheckNull(RezepteRec!AnzArtikel)
    If AnzRezeptArtikel% > 0 Then
    
      Unique$ = RezepteRec!Unique
      
      ArtikelRec.index = "Unique"
      For j% = 0 To (AnzRezeptArtikel% - 1)
          
          h$ = Unique$ + Format(j% + 1, "0")
          If j% >= 9 And j% <= 243 Then
              h$ = Unique$ + Chr$(j% + 65)   'für Privatrezepte
          End If
          
          ArtikelRec.Seek "=", h$
      
          If (ArtikelRec.NoMatch = False) Then
              ArtikelRec.Edit
              ArtikelRec!ImpErspart = 0
              If ArtikelRec!Imp = 2 Then
                  ArtikelRec!Imp = 1
                  SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(ArtikelRec!pzn)
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                  If Not TaxeRec.EOF Then
                    If Val(TaxeRec!OriginalPZN) > 0 Then
                      aktvk = TaxeRec!vk / 100#
                      SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                      If Not TaxeRec.EOF Then
                        originalvk = TaxeRec!vk / 100#
                        If aktvk > 0 And originalvk > 0 Then
                          If (originalvk - aktvk) >= 15 Or aktvk <= (originalvk * 0.85) Then
                            ArtikelRec!ImpErspart = (originalvk - aktvk)
                            RezepteRec!ImpErspart = dCheckNull(RezepteRec!ImpErspart) + (originalvk - aktvk)
                            ArtikelRec!Imp = 2
                          End If
                        End If
                      End If
                    End If
                  End If
              End If
              ArtikelRec.Update
          End If
      Next j%
      
      If dCheckNull(RezepteRec!ImpErspart) > 0 Then
      With AuswertungRec
        .index = "Unique"
        .Seek "=", RezepteRec!Kkasse, Left(RezepteRec!Unique, 4)
        If (.NoMatch) Then
            .AddNew
            AuswertungRec!Kkasse = RezepteRec!Kkasse
            AuswertungRec!Monat = Left(RezepteRec!Unique, 4)
            AuswertungRec!Rez_Gesamt = 0
            AuswertungRec!Rez_GesamtFAM = 0
            AuswertungRec!Rez_ImpFähig = 0
            AuswertungRec!Rez_ImpIst = 0
        Else
            .Edit
        End If
        AuswertungRec!ImpErspart = dCheckNull(AuswertungRec!ImpErspart) + dCheckNull(RezepteRec!ImpErspart)
        .Update
      End With
      End If
      RezepteRec.Update
    End If
  End If
  RezepteRec.MoveNext
Loop

        Call WinArtDebug("101")

Call DefErrPop
End Sub

Sub CreateAbrechnungsdaten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateAbrechnungsdaten")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim fld As Field
Dim i%, Jahr%

    Call WinArtDebug("400")
Set AbrechDatenTd = RezSpeicherDB.CreateTableDef("AbrechnungsDaten")

Set fld = AbrechDatenTd.CreateField("Unique", dbByte)
AbrechDatenTd.Fields.Append fld
Set fld = AbrechDatenTd.CreateField("Datum", dbText)
fld.Size = 6
AbrechDatenTd.Fields.Append fld

RezSpeicherDB.TableDefs.Append AbrechDatenTd

Set AbrechDatenRec = RezSpeicherDB.OpenRecordset("AbrechnungsDaten", dbOpenTable)
Jahr% = Format(Now, "YY")

For i% = 1 To 12
    AbrechDatenRec.AddNew
    AbrechDatenRec!Unique = i%
    If i% = 12 Then
        AbrechDatenRec!Datum = "0101" + Format(Jahr% + 1, "00")
    Else
        AbrechDatenRec!Datum = "01" + Format(i% + 1, "00") + Format(Jahr%, "00")
    End If
    AbrechDatenRec.Update
Next i%
    Call WinArtDebug("401")

Call DefErrPop
End Sub

Sub CreateAutIdemKrankenkassen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateAutIdemKrankenkassen")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim fld As Field
Dim i%
Dim AutIdemKkIdx As index
Dim ixFld As Field

    Call WinArtDebug("402")
Set AutIdemKkTd = RezSpeicherDB.CreateTableDef("AutIdemKrankenkassen")

Set fld = AutIdemKkTd.CreateField("Ik", dbLong)
AutIdemKkTd.Fields.Append fld

Set fld = AutIdemKkTd.CreateField("Name", dbText)
fld.Size = 48
AutIdemKkTd.Fields.Append fld

Set fld = AutIdemKkTd.CreateField("KbvNr", dbLong)
AutIdemKkTd.Fields.Append fld

Set fld = AutIdemKkTd.CreateField("AnzRezepte", dbLong)
AutIdemKkTd.Fields.Append fld

' Indizes für AutIdemKrankenkassen
Set AutIdemKkIdx = AutIdemKkTd.CreateIndex()
AutIdemKkIdx.Name = "Ik"
AutIdemKkIdx.Primary = True
AutIdemKkIdx.Unique = True
Set ixFld = AutIdemKkIdx.CreateField("Ik")
AutIdemKkIdx.Fields.Append ixFld
AutIdemKkTd.Indexes.Append AutIdemKkIdx

Set AutIdemKkIdx = AutIdemKkTd.CreateIndex()
AutIdemKkIdx.Name = "Name"
AutIdemKkIdx.Primary = False
AutIdemKkIdx.Unique = False
Set ixFld = AutIdemKkIdx.CreateField("Name")
AutIdemKkIdx.Fields.Append ixFld
AutIdemKkTd.Indexes.Append AutIdemKkIdx

RezSpeicherDB.TableDefs.Append AutIdemKkTd

Set AutIdemKkRec = RezSpeicherDB.OpenRecordset("AutIdemKrankenkassen", dbOpenTable)

    Call WinArtDebug("403")
Call DefErrPop
End Sub

Function CreateKassen%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateKassen%")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    Call WinArtDebug("300")
Dim KasseIdx As index
Dim KasseFld As Field
Dim ixFld As Field

Dim kk As Recordset

Dim kkey$
Dim ActRecNo&
Dim ret%

ret% = True

If Dir(KASSEN_DB) <> "" Then Kill KASSEN_DB
Set KassenDB = CreateDatabase(KASSEN_DB, dbLangGeneral) ', dbVersion30)

Set KasseTd = KassenDB.CreateTableDef("Kassen")
Set KasseFld = KasseTd.CreateField("Nummer", dbText)
KasseFld.Size = 9
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Kurz", dbText)
KasseFld.AllowZeroLength = False
KasseFld.Size = 10
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Name", dbText)
KasseFld.Size = 30
KasseFld.AllowZeroLength = False
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Strasse", dbText)
KasseFld.Size = 30
KasseFld.AllowZeroLength = True
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("PlzOrt", dbText)
KasseFld.Size = 30
KasseFld.AllowZeroLength = True
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Tel", dbText)
KasseFld.Size = 20
KasseFld.AllowZeroLength = True
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Fax", dbText)
KasseFld.Size = 20
KasseFld.AllowZeroLength = True
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Typ", dbByte)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Veraend", dbByte)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Zuzatyp", dbByte)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("ZuzaAbweichung", dbDouble)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Etikette", dbBoolean)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Anzeige", dbBoolean)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("ImportKZ", dbBoolean)
KasseTd.Fields.Append KasseFld
Set KasseFld = KasseTd.CreateField("Notiz", dbMemo)
KasseFld.AllowZeroLength = True
KasseTd.Fields.Append KasseFld

' Indizes für KASSE
Set KasseIdx = KasseTd.CreateIndex()
KasseIdx.Name = "Unique"
KasseIdx.Primary = True
KasseIdx.Unique = True
Set ixFld = KasseIdx.CreateField("Nummer")
KasseIdx.Fields.Append ixFld
KasseTd.Indexes.Append KasseIdx

Set KasseIdx = KasseTd.CreateIndex()
KasseIdx.Name = "Name"
KasseIdx.Primary = False
KasseIdx.Unique = False
Set ixFld = KasseIdx.CreateField("Name")
KasseIdx.Fields.Append ixFld
KasseTd.Indexes.Append KasseIdx

KassenDB.TableDefs.Append KasseTd

'Set kk = KassenDB.OpenRecordset("Kassen", dbOpenTable)
'
'FabsErrf% = Kkasse.IndexSearch(0, Space(9), FabsRecno&)
'If FabsErrf% = 13 Then FabsErrf% = 0
'Do While FabsErrf% = 0
'
'    Kkasse.GetRecord (FabsRecno& + 1)
'    kkey$ = Kkasse.Nummer
'    ActRecNo& = FabsRecno&
'    If Val(Kkasse.Nummer) > 0 Then
'      kk.AddNew
'      kk!Nummer = Kkasse.Nummer
'      kk!Name = RTrim(Kkasse.Name)
'      kk!kurz = RTrim(Kkasse.kurz)
'      kk!strasse = ""
'      kk!plzort = ""
'      kk!tel = ""
'      kk!fax = ""
'      kk!Typ = Kkasse.Typ
'      kk!Veraend = Abs((Kkasse.Veraend = "M"))
'      kk!zuzatyp = Kkasse.ZZTyp
'      kk!ZuzaAbweichung = Kkasse.zzab
'      kk!Etikette = (Kkasse.Eti = "J")
'      kk!Anzeige = Kkasse.Anzeige
'      kk!Notiz = ""
'      kk!strasse = " "
'      kk!plzort = " "
'      kk!tel = " "
'      kk!fax = " "
'      kk.Update
'    End If
'    FabsErrf% = Kkasse.IndexNext(0, ActRecNo&, kkey$, FabsRecno&)
'Loop
'kk.Close
'Set kk = Nothing

    Call WinArtDebug("301")
    
CreateKassen% = ret%
Call DefErrPop
End Function

Sub CreateParenteralTaxierungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateParenteralTaxierungen")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Td As TableDef
Dim fld As Field
Dim ix As Field
Dim idx As index

    Call WinArtDebug("404")
Set Td = RezSpeicherDB.CreateTableDef("ParenteralTaxierungen")

Set fld = Td.CreateField("Id", dbLong)
fld.Attributes = dbAutoIncrField
Td.Fields.Append fld
    
Set fld = Td.CreateField("RezepteUnique", dbText)
fld.Size = 14
Td.Fields.Append fld

Set fld = Td.CreateField("AnfMagInd", dbInteger)
Td.Fields.Append fld
    
'Set fld = Td.CreateField("PZN", dbText)
'fld.Size = 7
'fld.AllowZeroLength = True
Set fld = Td.CreateField("PZN", dbLong)
fld.DefaultValue = 0
Td.Fields.Append fld

Set fld = Td.CreateField("Text", dbText)
fld.Size = 50
Td.Fields.Append fld

Set fld = Td.CreateField("Menge", dbText)
fld.AllowZeroLength = True
fld.Size = 10
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("Einheit", dbText)
fld.AllowZeroLength = True
fld.Size = 2
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("Flag", dbByte)
Td.Fields.Append fld

Set fld = Td.CreateField("Kp", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("Gstufe", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("ActMenge", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("ActPreis", dbDouble)
Td.Fields.Append fld


' Indizes für ParenteralTaxierungen
Set idx = Td.CreateIndex()
idx.Name = "Id"
idx.Primary = True
idx.Unique = True
Set ix = idx.CreateField("Id")
idx.Fields.Append ix

Td.Indexes.Append idx

Set idx = Td.CreateIndex()
idx.Name = "RezepteUnique"
idx.Primary = False
idx.Unique = False
Set ix = idx.CreateField("RezepteUnique")
idx.Fields.Append ix
Set ix = idx.CreateField("AnfMagInd")
idx.Fields.Append ix

Td.Indexes.Append idx

RezSpeicherDB.TableDefs.Append Td

    Call WinArtDebug("405")
Call DefErrPop
End Sub

Sub CreateParenteralTaxmuster()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateParenteralTaxmuster")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Td As TableDef
Dim fld As Field
Dim ix As Field
Dim idx As index

    Call WinArtDebug("406")
Set Td = RezSpeicherDB.CreateTableDef("ParenteralTm")

Set fld = Td.CreateField("Id", dbLong)
fld.Attributes = dbAutoIncrField
Td.Fields.Append fld
    
Set fld = Td.CreateField("Bezeichnung", dbText)
fld.Size = 50
Td.Fields.Append fld

' Indizes für ParenteralTm
Set idx = Td.CreateIndex()
idx.Name = "Id"
idx.Primary = True
idx.Unique = True
Set ix = idx.CreateField("Id")
idx.Fields.Append ix

Td.Indexes.Append idx

RezSpeicherDB.TableDefs.Append Td


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set Td = RezSpeicherDB.CreateTableDef("ParenteralTmZeilen")

Set fld = Td.CreateField("Id", dbLong)
fld.Attributes = dbAutoIncrField
Td.Fields.Append fld
    
Set fld = Td.CreateField("TmId", dbLong)
Td.Fields.Append fld

Set fld = Td.CreateField("AnfMagInd", dbInteger)
Td.Fields.Append fld
    
'Set fld = Td.CreateField("PZN", dbText)
'fld.Size = 7
'fld.AllowZeroLength = True
Set fld = Td.CreateField("PZN", dbLong)
fld.DefaultValue = 0
Td.Fields.Append fld

Set fld = Td.CreateField("Text", dbText)
fld.Size = 50
Td.Fields.Append fld

Set fld = Td.CreateField("Menge", dbText)
fld.AllowZeroLength = True
fld.Size = 10
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("Einheit", dbText)
fld.AllowZeroLength = True
fld.Size = 2
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("Flag", dbByte)
Td.Fields.Append fld

Set fld = Td.CreateField("Kp", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("Gstufe", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("ActMenge", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("ActPreis", dbDouble)
Td.Fields.Append fld

Set fld = Td.CreateField("Packmittel", dbByte)
Td.Fields.Append fld

Set fld = Td.CreateField("AI", dbByte)
Td.Fields.Append fld

Set fld = Td.CreateField("WirkstoffMenge", dbDouble)
Td.Fields.Append fld

' Indizes für ParenteralTmZeilen
Set idx = Td.CreateIndex()
idx.Name = "Id"
idx.Primary = True
idx.Unique = True
Set ix = idx.CreateField("Id")
idx.Fields.Append ix

Td.Indexes.Append idx

Set idx = Td.CreateIndex()
idx.Name = "TmId"
idx.Primary = False
idx.Unique = False
Set ix = idx.CreateField("TmId")
idx.Fields.Append ix
Set ix = idx.CreateField("AnfMagInd")
idx.Fields.Append ix

Td.Indexes.Append idx

RezSpeicherDB.TableDefs.Append Td

    Call WinArtDebug("407")
Call DefErrPop
End Sub

Sub CreateAbrechnungsMeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateAbrechnungsMeldungen")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Td As TableDef
Dim fld As Field
Dim ix As Field
Dim idx As index

    Call WinArtDebug("408")
Set Td = RezSpeicherDB.CreateTableDef("AbrechnungsMeldungen")

Set fld = Td.CreateField("Id", dbLong)
fld.Attributes = dbAutoIncrField
Td.Fields.Append fld
    
Set fld = Td.CreateField("RezepteUnique", dbText)
fld.Size = 14
Td.Fields.Append fld

Set fld = Td.CreateField("LaufNr", dbInteger)
Td.Fields.Append fld
    
Set fld = Td.CreateField("fCode", dbInteger)
Td.Fields.Append fld
    
Set fld = Td.CreateField("fStatus", dbText)
fld.Size = 20
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("fKommentar", dbText)
fld.Size = 100
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("fWert", dbDouble)
Td.Fields.Append fld
    
Set fld = Td.CreateField("fristEnde", dbText)
fld.Size = 7
fld.AllowZeroLength = True
Td.Fields.Append fld

Set fld = Td.CreateField("fTCode", dbInteger)
Td.Fields.Append fld

Set fld = Td.CreateField("posNr", dbInteger)
Td.Fields.Append fld

Set fld = Td.CreateField("fKurzText", dbText)
fld.Size = 100
fld.AllowZeroLength = True
Td.Fields.Append fld

'Set fld = Td.CreateField("fLangText", dbText)
'fld.Size = 7
'fld.AllowZeroLength = True
'Td.Fields.Append fld

Set fld = Td.CreateField("fHauptfehler", dbByte)
Td.Fields.Append fld

Set fld = Td.CreateField("fVerbesserung", dbText)
fld.Size = 100
fld.AllowZeroLength = True
Td.Fields.Append fld


' Indizes für AbrechnungsMeldungen
Set idx = Td.CreateIndex()
idx.Name = "Id"
idx.Primary = True
idx.Unique = True
Set ix = idx.CreateField("Id")
idx.Fields.Append ix

Td.Indexes.Append idx

Set idx = Td.CreateIndex()
idx.Name = "RezepteUnique"
idx.Primary = False
idx.Unique = False
Set ix = idx.CreateField("RezepteUnique")
idx.Fields.Append ix
Set ix = idx.CreateField("LaufNr")
idx.Fields.Append ix

Td.Indexes.Append idx

RezSpeicherDB.TableDefs.Append Td

    Call WinArtDebug("409")
Call DefErrPop
End Sub

'Public Function dCheckNull(f As Field) As Double
Public Function dCheckNull(f As Variant) As Double
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("dCheckNull")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error Resume Next
If Not IsNull(f) Then
  dCheckNull = f
End If
Call DefErrPop
End Function

Public Function IstOriginal(pzn As Long, OriginalPZN As Long) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstOriginal")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim SQLStr$
Dim aktavp As Double, OrgAvp As Double


SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(pzn)
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If Not TaxeRec.EOF Then
  aktavp = TaxeRec!vk / 100#
  If aktavp > 0 Then
    If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
      IstOriginal = True
      OriginalPZN = TaxeRec!pzn
      OrgAvp = TaxeRec!vk / 100#
    ElseIf TaxeRec!OriginalPZN > 0 Then
      SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'      If Not TaxeRec.NoMatch Then
      If Not (TaxeRec.EOF) Then
        OrgAvp = TaxeRec!vk / 100#
        If aktavp > (OrgAvp - 15) And aktavp > (OrgAvp * 0.85) Then   'ist zwar kein Original, aber so teuer, dass es als solches gilt
          OriginalPZN = TaxeRec!pzn
          IstOriginal = True
        End If
      End If
    End If
    
    If IstOriginal Then
      'Suche nach Importen
      IstOriginal = False
      SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + CStr(OriginalPZN)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'      If Not TaxeRec.NoMatch Then
      If Not (TaxeRec.EOF) Then
        TaxeRec.MoveFirst
        Do
          If TaxeRec!pzn <> OriginalPZN And TaxeRec!pzn <> pzn Then
            aktavp = TaxeRec!vk / 100#
            If aktavp <= (OrgAvp - 15) Or aktavp <= (OrgAvp * 0.85) Then
              'es gibt also einen Import
              IstOriginal = True
              Exit Do
            End If
          End If
          TaxeRec.MoveNext
        Loop Until TaxeRec.EOF
      End If
    End If      'If IstOriginal Then
  End If    'If aktavp > 0 Then
End If

Call DefErrPop
End Function

Sub PruefeTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeTaetigkeiten")
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
Dim i%, k%, gef%, MindEiner%
Dim h$

DarfRezSpeicher = True
DarfImportKontrolle = True


For i% = 0 To (AnzTaetigkeiten% - 1)
    gef% = False
    MindEiner = (Taetigkeiten(i%).pers(0) > 0)
    For k% = 0 To 79
        If (Taetigkeiten(i%).pers(k%) = ActBenutzer%) Then
            gef% = True
            Exit For
        End If
    Next k%
    If (gef% = False) And (MindEiner) Then
        h = Trim(Taetigkeiten(i%).Taetigkeit)
        If (UCase(h) = UCase("Rezeptspeicher")) Then
            DarfRezSpeicher = False
        End If
        If (UCase(h) = UCase("ImportKontrolle")) Then
            DarfImportKontrolle = False
        End If
    End If
Next i%

Call DefErrPop

End Sub

Public Function sCheckNull(f As Field) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("sCheckNull")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error Resume Next
If Not IsNull(f) Then
  sCheckNull = f
End If
Call DefErrPop

End Function

Sub ScreenWerte(ScreenSizeHeight&, ScreenSizeWidth&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ScreenWerte")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim l&
Dim WorkRect As RectType

ScreenSizeWidth = (Screen.Width / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelX

ScreenSizeHeight = Screen.Height / Screen.TwipsPerPixelY
l = SystemParametersInfo(48, 0, WorkRect, 0)
ScreenSizeHeight = (WorkRect.Bottom - WorkRect.Top)
ScreenSizeHeight = ScreenSizeHeight * Screen.TwipsPerPixelY
Call DefErrPop

End Sub

Sub SelectAll(t As TextBox)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SelectAll")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
t.SelStart = 0
t.SelLength = Len(t.text)
Call DefErrPop
End Sub


Public Function ImportQuote#(Fam#, ImpFähig#, Optional Jahr$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImportQuote#")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ImpfProzent#
Dim iJahr%

iJahr% = Val(Left(Jahr, 2))
If iJahr% = 0 Then iJahr% = Val(Format(Now, "YY"))
If iJahr% = 4 And Val(Mid(Jahr, 3, 2)) < 6 Then   'gilt erst ab 1.6.04
  If Jahr > "" Or Month(Now) < 7 Then iJahr% = 3    'Abrechnung Juni kommt erst im Juli
End If
If Fam# > 0 Then
  ImpfProzent# = ImpFähig# * 100# / Fam#
End If
If ImpfProzent# <= 0 Then
    ImportQuote# = 0
ElseIf ImpfProzent# <= 5 Then
    If iJahr% >= 4 Then
        ImportQuote# = 0.8
    ElseIf iJahr% < 3 Then
        ImportQuote# = 0.9
    Else
        ImportQuote# = 1.2
    End If
ElseIf ImpfProzent# <= 10 Then      '17.5.02
    If iJahr% >= 4 Then
        ImportQuote# = 1.7
    ElseIf iJahr% < 3 Then
        ImportQuote# = 1.8
    Else
        ImportQuote# = 2.3
    End If
ElseIf ImpfProzent# <= 15 Then
    If iJahr% >= 4 Then
        ImportQuote# = 2.5
    ElseIf iJahr% < 3 Then
        ImportQuote# = 2.8
    Else
        ImportQuote# = 3.5
    End If
ElseIf ImpfProzent# <= 20 Then
    If iJahr% >= 4 Then
        ImportQuote# = 3.3
    ElseIf iJahr% < 3 Then
        ImportQuote# = 3.7
    Else
        ImportQuote# = 4.7
    End If
ElseIf ImpfProzent# <= 25 Then
    If iJahr% >= 4 Then
        ImportQuote# = 4.2
    ElseIf iJahr% < 3 Then
        ImportQuote# = 4.6
    Else
        ImportQuote# = 5.8
    End If
Else
    If iJahr% >= 4 Then
        ImportQuote# = 5#
    ElseIf iJahr% < 3 Then
        ImportQuote# = 5.5
    Else
        ImportQuote# = 7#
    End If
End If
Call DefErrPop
End Function


Function OpenRezeptSpeicher%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenRezeptSpeicher%")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%, i%, j%, ArtIndexDeb%
Dim s#
Dim tb As TableDef
Dim fld As Field
Dim RezeptIdx As index
Dim ArtikelIdx As index
Dim h$, t$, DebugStr$, SQLStr$
Dim l&
Dim ImpErspart As Boolean

ret% = False

Call WinArtDebug("OpenRezeptSpeicher")

Set Kkasse = New clsKkassen
'Kkasse.OpenDatei ("RW")

Call WinArtDebug("nach Kkasse.OpenDatei")

On Error Resume Next
For i% = 1 To 2
    Set RezSpeicherDB = OpenDatabase(REZ_SPEICHER, False, False)
    
    Call WinArtDebug("RezSPeicherDB: " + CStr(i) + Str(Err.Number))

    If ((Err = 3024) Or (Err = 3044)) Then
        On Error GoTo DefErr
        ret% = MessageBox("Rezeptspeicher noch nicht vorhanden!" + vbCr + "Soll er jetzt angelegt werden ?", vbYesNo Or vbQuestion)
        If (ret% = vbYes) Then
            ret% = CreateRezeptSpeicher%
        End If
        Exit For
    ElseIf (Err = 3343) Then
        On Error GoTo DefErr
        Call WinArtDebug("vor RepairDatabase")
        Call RepairDatabase(REZ_SPEICHER)
        Call WinArtDebug("nach RepairDatabase")
    ElseIf (Err = 0) Then
        ret% = True
        Exit For
    Else
        On Error GoTo DefErr
    End If
Next i%
On Error GoTo DefErr

Call WinArtDebug("nach RezSpeicherDB.open")

If (ret%) Then
    Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
    Set fld = RezepteRec.Fields("Knr")
    If (fld.Type = dbInteger) Then
        On Error Resume Next
        RezepteRec.Close
    '    SQLStr$ = "ALTER TABLE Rezepte MODIFY Knr INTEGER"
        SQLStr$ = "ALTER TABLE Rezepte ALTER COLUMN Knr LONG"
        RezSpeicherDB.Execute SQLStr$
    End If

    Call WinArtDebug("1")

    Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
    Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
    Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)
    
    Call WinArtDebug("2")

    On Error Resume Next
    Set AbrechDatenRec = RezSpeicherDB.OpenRecordset("AbrechnungsDaten", dbOpenTable)
    If Err.Number = 3011 Then
        Call CreateAbrechnungsdaten
    End If
    Err.Clear
    Call WinArtDebug("3")

    Set AutIdemKkRec = RezSpeicherDB.OpenRecordset("AutIdemKrankenkassen", dbOpenTable)
    If Err.Number = 3011 Then
        Call CreateAutIdemKrankenkassen
    End If
    Err.Clear
    Call WinArtDebug("4")

    Set ParenteralTaxierungRec = RezSpeicherDB.OpenRecordset("ParenteralTaxierungen", dbOpenTable)
    If Err.Number = 3011 Then
        Call CreateParenteralTaxierungen
    End If
    Err.Clear
    Call WinArtDebug("5")

    s# = dCheckNull(ParenteralTaxierungRec!Verwurf)
    If Err.Number = 3265 Then
      ParenteralTaxierungRec.Close
      Err.Clear
      Set tb = RezSpeicherDB.TableDefs("ParenteralTaxierungen")
      
      tb.Fields.Append tb.CreateField("Verwurf", dbByte)
      
        Set ParenteralTaxierungRec = RezSpeicherDB.OpenRecordset("ParenteralTaxierungen", dbOpenTable)
    End If
    Err.Clear
    Call WinArtDebug("5a")
    
    
    Set ParenteralTmRec = RezSpeicherDB.OpenRecordset("ParenteralTM", dbOpenTable)
    If Err.Number = 3011 Then
        Call CreateParenteralTaxmuster
    End If
    Err.Clear
    Call WinArtDebug("6")

    Set AbrechnungsMeldungenRec = RezSpeicherDB.OpenRecordset("AbrechnungsMeldungen", dbOpenTable)
    If Err.Number = 3011 Then
        Call CreateAbrechnungsMeldungen
    End If
    Err.Clear
    Call WinArtDebug("7")

    If (UCase(App.EXEName) <> "WINREZDR") Then
        RezepteRec.MoveFirst
        s# = dCheckNull(RezepteRec!Fam)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("FAM", dbDouble)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("8")
    
        
        s# = dCheckNull(RezepteRec!ImpFähig)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("ImpFähig", dbDouble)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("9")
    
        
        s# = dCheckNull(RezepteRec!ImpIst)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("ImpIst", dbDouble)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("10")
    
        
        s# = dCheckNull(RezepteRec!knr)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("Knr", dbInteger)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("11")
    
    
        t$ = sCheckNull(RezepteRec!Arzt)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("Arzt", dbText)
          fld.Size = 10
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("12")
    
    
        s# = dCheckNull(RezepteRec!AbgabeDatum)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("AbgabeDatum", dbDate)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("13")
    
        
        t$ = sCheckNull(RezepteRec!AbrechnungsMonat)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("AbrechnungsMonat", dbText)
          fld.Size = 4
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("14")
    
    
        s# = dCheckNull(RezepteRec!DruckZeit)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("DruckZeit", dbInteger)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("15")
    
    
        s# = dCheckNull(RezepteRec!RabattWert)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("RabattWert", dbDouble)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("16")
    
        
        s# = dCheckNull(RezepteRec!ImpErspart)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          tb.Fields.Append tb.CreateField("ImpErspart", dbDouble)
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
          ImpErspart = True
        End If
        Err.Clear
        Call WinArtDebug("17")
    
        
        s# = dCheckNull(RezepteRec!AvpRezeptNr)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("AvpRezeptNr", dbText)
          fld.Size = 11
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          tb.Fields.Append tb.CreateField("AvpLaufNr", dbLong)
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("18")
    
        
        s# = dCheckNull(RezepteRec!Verfügbarkeit)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          
          Set fld = tb.CreateField("Verfügbarkeit", dbText)
          fld.Size = 3
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("19")
    
        
        s# = dCheckNull(RezepteRec!kKassenIk)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          
          tb.Fields.Append tb.CreateField("kKassenIk", dbLong)
          
    '      Set fld = tb.CreateField("kKassenName", dbText)
    '      fld.Size = 48
    '      fld.AllowZeroLength = True
    '      tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("20")
    
        
        s# = dCheckNull(RezepteRec!TransaktionsID)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("TransaktionsID", dbLong)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("21")
    
        
        s# = dCheckNull(RezepteRec!ParenteralHash)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("ParenteralHash", dbText)
          fld.Size = 40
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("22")
    
        
        s# = dCheckNull(RezepteRec!AbrechnungsStatus)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("AbrechnungsStatus", dbByte)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("23")
    
        
        s# = dCheckNull(RezepteRec!AbrechnungsMeldung)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          
          Set fld = tb.CreateField("AbrechnungsMeldung", dbText)
          fld.Size = 40
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("24")
    
        
        s# = dCheckNull(RezepteRec!rzieferId)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          
          Set fld = tb.CreateField("rzLieferId", dbText)
          fld.Size = 40
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("25")
    
        
        s# = dCheckNull(RezepteRec!SendeStatus)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("SendeStatus", dbByte)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("26")
    
        
        s# = dCheckNull(RezepteRec!P302)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("P302", dbByte)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27")
    
        
        s# = dCheckNull(RezepteRec!HerstRabatt)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("HerstRabatt", dbDouble)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27a")
    
        s# = dCheckNull(RezepteRec!pCharge)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("pCharge", dbText)
          fld.Size = 30
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27b")
    
        s# = dCheckNull(RezepteRec!HashErstellungsDatum)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("HashErstellungsDatum", dbText)
          fld.Size = 13
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27c")
    
        s# = dCheckNull(RezepteRec!VebNr)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("VebNr", dbLong)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27d")
    
        s# = dCheckNull(RezepteRec!PauschaleNr)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("PauschaleNr", dbLong)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27e")
    
        s# = dCheckNull(RezepteRec!VerkLiRe)
        If Err.Number = 3265 Then
          RezepteRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Rezepte")
          Set fld = tb.CreateField("VerkLiRe", dbByte)
          tb.Fields.Append fld
          Set fld = tb.CreateField("AutIdemKbvNr", dbLong)
          tb.Fields.Append fld
          Set fld = tb.CreateField("AVWGik", dbLong)
          tb.Fields.Append fld
          
          Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte")
        End If
        Err.Clear
        Call WinArtDebug("27f")
    
        
        ArtikelRec.MoveFirst
        s# = dCheckNull(ArtikelRec!ImpErspart)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("ImpErspart", dbDouble)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("28")
    
    
        s# = dCheckNull(ArtikelRec!Warenzeichen)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("Warenzeichen", dbByte)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("29")
    
        
        s# = dCheckNull(ArtikelRec!HmNummer)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          Set fld = tb.CreateField("HmNummer", dbText)
          fld.Size = 10
          tb.Fields.Append fld
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("30")
    
    
        s# = dCheckNull(ArtikelRec!HmFaktor)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("HmFaktor", dbInteger)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("31")
    
    
        s# = dCheckNull(ArtikelRec!HmStückPreis)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("HmStückPreis", dbDouble)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("32")
    
        s# = dCheckNull(ArtikelRec!HerstRabattPrivat130Brutto)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("HerstRabattPrivat130Brutto", dbDouble)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("32a")
    
        s# = dCheckNull(ArtikelRec!TkkPznDruck)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("TkkPznDruck", dbByte)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("32b")
    
        s# = dCheckNull(ArtikelRec!AutIdemKreuz)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("AutIdemKreuz", dbByte)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("32c")
    
        s# = dCheckNull(ArtikelRec!Fam)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          tb.Fields.Append tb.CreateField("Fam", dbByte)
          tb.Fields.Append tb.CreateField("ePreis", dbDouble)
          tb.Fields.Append tb.CreateField("AutIdem", dbByte)
          tb.Fields.Append tb.CreateField("N0", dbByte)
          tb.Fields.Append tb.CreateField("MagIndex", dbByte)
          tb.Fields.Append tb.CreateField("Appli", dbByte)
          
          Set fld = tb.CreateField("zusatz1", dbText)
          fld.Size = 36
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          Set fld = tb.CreateField("zusatz2", dbText)
          fld.Size = 50 ' 36
          fld.AllowZeroLength = True
          tb.Fields.Append fld
    
          tb.Fields.Append tb.CreateField("NichtInTaxe", dbByte)
          tb.Fields.Append tb.CreateField("IstWg4", dbByte)
          tb.Fields.Append tb.CreateField("PlusMehrKosten", dbByte)
          tb.Fields.Append tb.CreateField("ZuzahlungsErlass", dbByte)
          tb.Fields.Append tb.CreateField("MietDauer", dbInteger)
          tb.Fields.Append tb.CreateField("VdbPauschale", dbByte)
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("32d")
    
        s# = dCheckNull(ArtikelRec!VebNr)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          Set fld = tb.CreateField("VebNr", dbLong)
          tb.Fields.Append fld
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("27d2")
    
        s# = dCheckNull(ArtikelRec!PauschaleNr)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          Set fld = tb.CreateField("PauschaleNr", dbLong)
          tb.Fields.Append fld
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("27e")
    
        s# = dCheckNull(ArtikelRec!Fortsetzung)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          tb.Fields.Append tb.CreateField("Fortsetzung", dbByte)
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("27f")
    
        s# = dCheckNull(ArtikelRec!HmAbrechnungsKz)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          Set fld = tb.CreateField("HmAbrechnungsKz", dbText)
          fld.Size = 2
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
        Call WinArtDebug("27ff")
    
          
        s# = dCheckNull(ArtikelRec!zusatz3)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          Set fld = tb.CreateField("zusatz3", dbText)
          fld.Size = 50 ' 36
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
          
        s# = dCheckNull(ArtikelRec!AuseinzelungBtm)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          tb.Fields.Append tb.CreateField("AuseinzelungBtm", dbByte)
          
'            SQLStr$ = "ALTER TABLE Artikel ALTER COLUMN Zusatz2 TEXT(50)"
'            RezSpeicherDB.Execute SQLStr$
'            SQLStr$ = "ALTER TABLE Artikel ALTER COLUMN Zusatz3 TEXT(50)"
'            RezSpeicherDB.Execute SQLStr$
          
            Set fld = tb.CreateField("Zusatz2a", dbText)
            fld.Size = 10
            fld.AllowZeroLength = True
            tb.Fields.Append fld
            Set fld = tb.CreateField("Zusatz3a", dbText)
            fld.Size = 10
            fld.AllowZeroLength = True
            tb.Fields.Append fld
          
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
          
        s# = dCheckNull(ArtikelRec!LEGS)
        If Err.Number = 3265 Then
          ArtikelRec.Close
          Err.Clear
          
          Set tb = RezSpeicherDB.TableDefs("Artikel")
          
          Set fld = tb.CreateField("LEGS", dbText)
          fld.Size = 7
          fld.AllowZeroLength = True
          tb.Fields.Append fld
          
          Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel")
        End If
        Err.Clear
          
        
        AuswertungRec.MoveFirst
        s# = dCheckNull(AuswertungRec!abr_gesamtFAM)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("Abr_GesamtFAM", dbDouble)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("33")
    
        
        s# = dCheckNull(AuswertungRec!Saldo)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("Saldo", dbDouble)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("34")
    
        
        s# = dCheckNull(AuswertungRec!GutHaben)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("Guthaben", dbDouble)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("35")
    
        s# = dCheckNull(AuswertungRec!RezAnzahl)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("RezAnzahl", dbLong)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("36")
    
        s# = dCheckNull(AuswertungRec!ImpErspart)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("ImpErspart", dbDouble)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("37")
    
        s# = dCheckNull(AuswertungRec!abr_ImpErspart)
        If Err.Number = 3265 Then
          AuswertungRec.Close
          Err.Clear
          Set tb = RezSpeicherDB.TableDefs("Auswertung")
          tb.Fields.Append tb.CreateField("abr_ImpErspart", dbDouble)
          Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung")
        End If
        Err.Clear
        Call WinArtDebug("38")
    
       
        RezepteRec.index = "Unique"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
            
             ' Indizes für REZEPTE
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "Unique"
            RezeptIdx.Primary = True
            RezeptIdx.Unique = True
            Set fld = RezeptIdx.CreateField("Unique")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("39")
    
        
        RezepteRec.index = "Kasse"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "Kasse"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("Kkasse")
            RezeptIdx.Fields.Append fld
            Set fld = RezeptIdx.CreateField("VerkDatum")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("40")
    
        RezepteRec.index = "RezeptNr"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
            
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "RezeptNr"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("RezeptNR")
            RezeptIdx.Fields.Append fld
            Set fld = RezeptIdx.CreateField("VerkDatum")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("41")
    
        
        RezepteRec.index = "KasseDruck"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "KasseDruck"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("Kkasse")
            RezeptIdx.Fields.Append fld
            Set fld = RezeptIdx.CreateField("DruckDatum")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("42")
    
        
        RezepteRec.index = "AvpRezeptNr"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "AvpRezeptNr"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("AvpRezeptNr")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("43")
    
        
        RezepteRec.index = "AvpLaufNr"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "AvpLaufNr"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("AvpLaufNr")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("44")
    
        
        RezepteRec.index = "TransaktionsID"
    '    If Err.Number = 3015 Then
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "TransaktionsID"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("TransaktionsID")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Err.Clear
        Call WinArtDebug("45")
    
        RezepteRec.index = "DruckDatum"
        If (Err.Number) Then
            RezepteRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Rezepte")
        
            Set RezeptIdx = tb.CreateIndex()
            RezeptIdx.Name = "DruckDatum"
            RezeptIdx.Primary = False
            RezeptIdx.Unique = False
            
            Set fld = RezeptIdx.CreateField("DruckDatum")
            RezeptIdx.Fields.Append fld
            Set fld = RezeptIdx.CreateField("Unique")
            RezeptIdx.Fields.Append fld
            tb.Indexes.Append RezeptIdx
            Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        End If
        Call WinArtDebug("46")
    
        
            
        If (ArtIndexDebug%) Then
            Call WinArtDebug("47")
    
            ArtikelRec.Close
            Err.Clear
            Set tb = RezSpeicherDB.TableDefs("Artikel")
        
            tb.Indexes.Delete "Unique"
            DebugStr$ = DebugStr + "Nach Index-Delete:" + Str$(Err.Number) + vbCrLf
            Err.Clear
        
            Call WinArtDebug("48")
    
            DebugStr$ = DebugStr$ + "Vor Index-Erzeugen:" + Str$(Err.Number) + vbCrLf
            
            Set ArtikelIdx = tb.CreateIndex()
            ArtikelIdx.Name = "Unique"
            ArtikelIdx.Primary = True
            ArtikelIdx.Unique = False   ' True
            
            DebugStr$ = DebugStr$ + "Vor IndexFeld-Erzeugen:" + Str$(Err.Number) + vbCrLf
            
            Set fld = ArtikelIdx.CreateField("Unique")
            ArtikelIdx.Fields.Append fld
            tb.Indexes.Append ArtikelIdx
            
            DebugStr$ = DebugStr$ + "Nach Index-Erzeugen:" + Str$(Err.Number) + vbCrLf
            
            Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
            
            Call WinArtDebug("49")
    
            ArtIndexDeb% = FreeFile
            Open "ArtIndex.deb" For Output Access Write Shared As #ArtIndexDeb%
            Print #ArtIndexDeb%, , DebugStr$
            Close #ArtIndexDeb%
        End If
    End If
End If

    Call WinArtDebug("50")

If ImpErspart Then Call AddImpErspart
    Call WinArtDebug("200")

On Error Resume Next
Set KassenDB = OpenDatabase(KASSEN_DB, False, False)
If ((Err = 3024) Or (Err = 3044)) Then
    ret% = CreateKassen
End If
Set KasseRec = KassenDB.OpenRecordset("Kassen", dbOpenTable)

s# = dCheckNull(KasseRec!VebNr)
If Err.Number = 3265 Then
  KasseRec.Close
  Err.Clear
  Set tb = KassenDB.TableDefs("Kassen")
  tb.Fields.Append tb.CreateField("VebNr", dbLong)
  Set KasseRec = KassenDB.OpenRecordset("Kassen", dbOpenTable)
End If
Err.Clear
    
    Call WinArtDebug("202")
On Error GoTo DefErr
Call PruefeTaetigkeiten
    Call WinArtDebug("203")

Call AbrechMonatErmitteln
    Call WinArtDebug("204")
    
If (para.MARS) Then
    Call WinArtDebug("Mars_OpenRezeptSpeicher")
    
    On Error Resume Next
    For i% = 1 To 2
        Set MarsRezSpeicherDB = OpenDatabase(MARS_REZ_SPEICHER, False, False)
        
        Call WinArtDebug("MarsRezSpeicherDB: " + CStr(i) + Str(Err.Number))
    
        If ((Err = 3024) Or (Err = 3044)) Then
            On Error GoTo DefErr
            ret% = MessageBox("MARS-Rezeptspeicher noch nicht vorhanden!" + vbCr, vbCritical)
'            If (ret% = vbYes) Then
'                ret% = CreateRezeptSpeicher%
'            End If
            Exit For
        ElseIf (Err = 3343) Then
            On Error GoTo DefErr
            Call WinArtDebug("vor Mars_RepairDatabase")
            Call RepairDatabase(MARS_REZ_SPEICHER)
            Call WinArtDebug("nach Mars_RepairDatabase")
        ElseIf (Err = 0) Then
            ret% = True
            Exit For
        Else
            On Error GoTo DefErr
        End If
    Next i%
    On Error GoTo DefErr
    
    Call WinArtDebug("nach MarsRezSpeicherDB.open")
    
    If (ret%) Then
        Set MarsRezepteRec = MarsRezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        Set MarsArtikelRec = MarsRezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
    End If
End If


Call WinArtDebug("OpenRezeptSpeicher ENDE")

OpenRezeptSpeicher% = ret%

Call DefErrPop
End Function

Function CreateRezeptSpeicher%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateRezeptSpeicher%")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%

Dim ixFld As Field

Dim RezepteFld As Field
Dim ArtikelFld As Field
Dim AuswertungFld As Field

Dim ArtikelIdx As index
Dim RezeptIdx As index
Dim AuswertungIdx As index

ret% = True
 
If Dir(REZ_SPEICHER) <> "" Then Kill REZ_SPEICHER
Set RezSpeicherDB = CreateDatabase(REZ_SPEICHER, dbLangGeneral) ', dbVersion30)


' Tabelle REZEPTE
Set RezepteTd = RezSpeicherDB.CreateTableDef("Rezepte")

Set RezepteFld = RezepteTd.CreateField("Unique", dbText)
RezepteFld.Size = 14
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("RezeptNr", dbText)
RezepteFld.Size = 13
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("Kkasse", dbText)
RezepteFld.Size = 9
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("GebFrei", dbBoolean)
RezepteTd.Fields.Append RezepteFld


Set RezepteFld = RezepteTd.CreateField("RezGebSumme", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("RezSumme", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("RabattWert", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("FAM", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("ImpFähig", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("ImpIst", dbDouble)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("AnzArtikel", dbByte)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("Bundesland", dbInteger)
RezepteTd.Fields.Append RezepteFld
Set RezepteFld = RezepteTd.CreateField("Kasse", dbInteger)
RezepteTd.Fields.Append RezepteFld
Set RezepteFld = RezepteTd.CreateField("Verordnung", dbInteger)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("VerkDatum", dbText)
RezepteFld.Size = 6
RezepteTd.Fields.Append RezepteFld
Set RezepteFld = RezepteTd.CreateField("VerkZeit", dbInteger)
RezepteTd.Fields.Append RezepteFld
Set RezepteFld = RezepteTd.CreateField("VerkComp", dbText)
RezepteFld.Size = 2
RezepteTd.Fields.Append RezepteFld
Set RezepteFld = RezepteTd.CreateField("PersonalNr", dbByte)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("DruckDatum", dbText)
RezepteFld.Size = 6
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("InstitutsKz", dbText)
RezepteFld.Size = 7
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("Knr", dbInteger)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("Arzt", dbText)
RezepteFld.Size = 10
RezepteFld.AllowZeroLength = True
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("AbgabeDatum", dbDate)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("AbrechnungsMonat", dbText)
RezepteFld.Size = 4
RezepteFld.AllowZeroLength = True
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("DruckZeit", dbInteger)
RezepteTd.Fields.Append RezepteFld

Set RezepteFld = RezepteTd.CreateField("ImpErspart", dbDouble)
RezepteTd.Fields.Append RezepteFld


' Indizes für REZEPTE
Set RezeptIdx = RezepteTd.CreateIndex()
RezeptIdx.Name = "Unique"
RezeptIdx.Primary = True
RezeptIdx.Unique = True
Set ixFld = RezeptIdx.CreateField("Unique")
RezeptIdx.Fields.Append ixFld
RezepteTd.Indexes.Append RezeptIdx

Set RezeptIdx = RezepteTd.CreateIndex()
RezeptIdx.Name = "Kasse"
RezeptIdx.Primary = False
RezeptIdx.Unique = False
Set ixFld = RezeptIdx.CreateField("Kkasse")
RezeptIdx.Fields.Append ixFld
Set ixFld = RezeptIdx.CreateField("VerkDatum")
RezeptIdx.Fields.Append ixFld
RezepteTd.Indexes.Append RezeptIdx

Set RezeptIdx = RezepteTd.CreateIndex()
RezeptIdx.Name = "RezeptNr"
RezeptIdx.Primary = False
RezeptIdx.Unique = False
Set ixFld = RezeptIdx.CreateField("RezeptNr")
RezeptIdx.Fields.Append ixFld
Set ixFld = RezeptIdx.CreateField("VerkDatum")
RezeptIdx.Fields.Append ixFld
RezepteTd.Indexes.Append RezeptIdx

Set RezeptIdx = RezepteTd.CreateIndex()
RezeptIdx.Name = "KasseDruck"
RezeptIdx.Primary = False
RezeptIdx.Unique = False
Set ixFld = RezeptIdx.CreateField("Kkasse")
RezeptIdx.Fields.Append ixFld
Set ixFld = RezeptIdx.CreateField("DruckDatum")
RezeptIdx.Fields.Append ixFld
RezepteTd.Indexes.Append RezeptIdx

RezSpeicherDB.TableDefs.Append RezepteTd


' Tabelle ARTIKEL
Set ArtikelTd = RezSpeicherDB.CreateTableDef("Artikel")

Set ArtikelFld = ArtikelTd.CreateField("Unique", dbText)
ArtikelFld.Size = 15
ArtikelTd.Fields.Append ArtikelFld

'Set ArtikelFld = ArtikelTd.CreateField("Pzn", dbText)
'ArtikelFld.Size = 7
Set ArtikelFld = ArtikelTd.CreateField("Pzn", dbLong)
ArtikelFld.DefaultValue = 0
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Text", dbText)
ArtikelFld.Size = 36
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Preis", dbDouble)
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Flag", dbByte)
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Zuz", dbSingle)
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Imp", dbByte)
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("Faktor", dbInteger)
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("ScreenPzn", dbText)
ArtikelFld.Size = 10
'Set ArtikelFld = ArtikelTd.CreateField("ScreenPzn", dbLong)
'ArtikelFld.DefaultValue = 0
ArtikelTd.Fields.Append ArtikelFld
Set ArtikelFld = ArtikelTd.CreateField("ImpErspart", dbDouble)
ArtikelTd.Fields.Append ArtikelFld


' Indizes für ARTIKEL
Set ArtikelIdx = ArtikelTd.CreateIndex()
ArtikelIdx.Name = "Unique"
ArtikelIdx.Primary = True
ArtikelIdx.Unique = True
Set ixFld = ArtikelIdx.CreateField("Unique")
ArtikelIdx.Fields.Append ixFld
ArtikelTd.Indexes.Append ArtikelIdx

RezSpeicherDB.TableDefs.Append ArtikelTd




' Tabelle AUSWERTUNG
Set AuswertungTd = RezSpeicherDB.CreateTableDef("Auswertung")
Set AuswertungFld = AuswertungTd.CreateField("Kkasse", dbText)
AuswertungFld.Size = 9
AuswertungTd.Fields.Append AuswertungFld

Set AuswertungFld = AuswertungTd.CreateField("Monat", dbText)
AuswertungFld.Size = 4
AuswertungTd.Fields.Append AuswertungFld

Set AuswertungFld = AuswertungTd.CreateField("Abr_Gesamt", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Abr_GesamtFAM", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Abr_ImpFähig", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Abr_ImpIst", dbDouble)
AuswertungTd.Fields.Append AuswertungFld

Set AuswertungFld = AuswertungTd.CreateField("Rez_Gesamt", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Rez_GesamtFAM", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Rez_ImpFähig", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Rez_ImpIst", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("Abr_Guthaben", dbDouble)
AuswertungTd.Fields.Append AuswertungFld
Set AuswertungFld = AuswertungTd.CreateField("RezAnzahl", dbLong)
AuswertungTd.Fields.Append AuswertungFld

Set AuswertungFld = AuswertungTd.CreateField("Saldo", dbDouble)
AuswertungTd.Fields.Append AuswertungFld


' Indizes für AUSWERTUNG
Set AuswertungIdx = AuswertungTd.CreateIndex()
AuswertungIdx.Name = "Unique"
AuswertungIdx.Primary = True
AuswertungIdx.Unique = True
Set ixFld = AuswertungIdx.CreateField("Kkasse")
AuswertungIdx.Fields.Append ixFld
Set ixFld = AuswertungIdx.CreateField("Monat")
AuswertungIdx.Fields.Append ixFld
AuswertungTd.Indexes.Append AuswertungIdx

RezSpeicherDB.TableDefs.Append AuswertungTd

CreateRezeptSpeicher% = ret%

Call DefErrPop
End Function

Public Function xVal(ByVal x As String) As Double
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("xVal")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim i As Integer
Dim j As Integer

i = InStr(x, ",")
j = InStr(x, ".")

If i > 0 Then
  If j = 0 Then
    Mid(x, i, 1) = "."
  Else
    'erst die Tausendertrennzeichen weg
    While j > 0
      x = Left(x, j - 1) + Mid(x, j + 1)
      j = InStr(x, ".")
    Wend
    i = InStr(x, ",")
    Mid(x, i, 1) = "."
  End If
End If
xVal = Val(x)
Call DefErrPop
End Function

Function CheckVerfügbarkeit$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckVerfügbarkeit$")
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
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%
Dim ret$, h2$

ret$ = ""
For i% = 0 To 2
    h2$ = Trim(frmAction.cboVerfügbarkeit(i%).text)
    If (h2$ = "") Then
        ret$ = ret$ + "1"
    Else
        ret$ = ret$ + Left$(h2$, 1)
    End If
Next i%

CheckVerfügbarkeit$ = ret$

Call DefErrPop
End Function


