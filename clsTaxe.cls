VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaxe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim dKurz$
Dim dAnzFields%

Dim fKurz$()
Dim fBezeichnung$()

Private Const DefErrModul = "clsTaxe.cls"

Private Sub Class_Initialize()

Set Taxe1 = Me

End Sub

Function OpenDatenbank(DBName$, Optionen%, Schreibgeschuetzt%) As Database
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatenbank")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
'Call clsError.DefErrAbort
On Error GoTo 0
Err.Raise 999 + vbObjectError, "abc", "Fehler in OpenDatenbank"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, erg%

Set TaxeDB = OpenDatabase(DBName$, Optionen%, Schreibgeschuetzt%)
'If (erg% > 0) Then
'    erg% = MsgBox("Die Windows-Taxe existiert auf Laufwerk " + TaxeLw$ + ": nicht !", vbInformation)
'End If

If (Err = 0) Then
    dKurz$ = "TAXE"
    dAnzFields% = TaxeDB.TableDefs("TAXE").Fields.Count
    ReDim fBezeichnung$(dAnzFields%)
    For i% = 0 To (dAnzFields% - 1)
        fBezeichnung$(i%) = TaxeDB.TableDefs("TAXE").Fields(i%).Name
    Next i%
End If

Set OpenDatenbank = TaxeDB

Call clsError.DefErrPop
End Function

Function TaxeValue$(FeldName$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("HoleTaxeInfoWert$")
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
Dim ret$, h$
Dim fldRecordset As Field

ret$ = ""

Set fldRecordset = TaxeRec.Fields(FeldName$)
Select Case UCase(FeldName$)
    Case "PZN"
        ret$ = Format(fldRecordset.Value, "0000000")
        ret$ = " " + ret$
    Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT"
        ret$ = Format(fldRecordset.Value / 100#, "0.00")
        If (ret$ = "0.00") Or (ret$ = "0,00") Then
            ret$ = "0"
        End If
        ret$ = " " + ret$
    Case Else
        ret$ = " " + LTrim$(fldRecordset.Value)
End Select

TaxeValue$ = ret$

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%

'Call DefErrPop

End Function

Function FeldBezeichnung$(iFeld As Variant)
Dim ind%
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
'    ind% = SucheFeldInd%(iFeld)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldBezeichnung$ = ret$

End Function

Function FeldKurz$(iFeld As Variant)
Dim ind%
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
'    ind% = SucheFeldInd%(iFeld)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldKurz$ = ret$

End Function

Public Property Get DateiBezeichnung$()
DateiBezeichnung$ = "Taxe"
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property

