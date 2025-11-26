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


Private Const DefErrModul = "TAXE.CLS"

Private Sub Class_Initialize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Class_Initialize")
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
Set Taxe1 = Me

Call clsError.DefErrPop
End Sub

Function OpenDatenbank(DBName$, Optionen%, Schreibgeschuetzt%) As Database
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatenbank")
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
Dim i%, erg%

Set TaxeDB = OpenDatabase(DBName$, Optionen%, Schreibgeschuetzt%)
'If (erg% > 0) Then
'    erg% = MsgBox("Die Windows-Taxe existiert auf Laufwerk " + TaxeLw$ + ": nicht !", vbInformation)
'End If

If (Err = 0) Then
    dKurz$ = "TAXE"
    dAnzFields% = TaxeDB.TableDefs("TAXE").Fields.Count + 1
    ReDim fBezeichnung$(dAnzFields% + 1)
    fBezeichnung$(0) = "ImportKz"
    For i% = 0 To (dAnzFields% - 2)
        fBezeichnung$(i% + 1) = TaxeDB.TableDefs("TAXE").Fields(i%).Name
    Next i%
End If

Set OpenDatenbank = TaxeDB

Call clsError.DefErrPop
End Function

Function TaxeValue$(FeldName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TaxeValue$")
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
Dim ret$, h$, iFeldName$
Dim fldRecordset As Field

ret$ = ""

iFeldName$ = UCase(FeldName$)
If (iFeldName$ = "IMPORTKZ") Then
    If (TaxeRec!OriginalPZN > 0) Then
        If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
            ret$ = "i"
        Else
            ret$ = "I"
        End If
    End If
Else
    Set fldRecordset = TaxeRec.Fields(FeldName$)
    Select Case iFeldName$
        Case "PZN"
            ret$ = Format(fldRecordset.Value, "0000000")
        Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT", "HAP"
            ret$ = Format(fldRecordset.Value / 100#, "0.00")
            If (ret$ = "0.00") Or (ret$ = "0,00") Then
                ret$ = "0"
            End If
        Case "HAPSPANNE"
            ret$ = Format(fldRecordset.Value / 10#, "0.0")
            If (ret$ = "0.0") Or (ret$ = "0,0") Then
                ret$ = "0"
            End If
        Case "ZUZAHLUNG"
            ret$ = MachTaxeAuswahlZuzahlung$(TaxeRec)
        Case "GUELTIG", "LPREISAENDERUNG"
            If (iFeldName$ = "GUELTIG") Then
                h$ = LTrim$(fldRecordset.Value)
            Else
                h$ = Right$(String(6, "0") + Format(fldRecordset.Value, "0"), 6)
            End If
            ret$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "."
            If (Val(Right$(h$, 2)) < 50) Then
                ret$ = ret$ + "20"
            Else
                ret$ = ret$ + "19"
            End If
            ret$ = ret$ + Right$(h$, 2)
        Case Else
            ret$ = LTrim$(fldRecordset.Value)
    End Select
End If

TaxeValue$ = " " + ret$

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%

Call clsError.DefErrPop
End Function

Function FeldBezeichnung$(iFeld As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldBezeichnung$")
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

Call clsError.DefErrPop
End Function

Function FeldKurz$(iFeld As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldKurz$")
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

Call clsError.DefErrPop
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

Function MachTaxeAuswahlZuzahlung$(tRec As Recordset)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MachTaxeAuswahlZuzahlung$")
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
Dim ch$, ret$, ZuzWert!, i%, zz%, VglWert!, AVP!

ret$ = ""

If (tRec!ZuzFrei = 0) And (tRec!ErstattungsFähig < 3) Then
    ZuzWert! = tRec!ZuzaNeu
        
    AVP! = tRec!VK  '/ 100
    If ((tRec!FestKz) And (tRec!FESTBETRAG < AVP!)) Then
        VglWert! = tRec!FESTBETRAG
    Else
        VglWert! = AVP!
    End If
        
    If ((VglWert! > 0!) And (ZuzWert! > VglWert!)) Then
        ret$ = "VK"
    ElseIf (ZuzWert! > 0!) Then
        ret$ = Format$(ZuzWert! / 100, "#0.00")
    End If
End If

MachTaxeAuswahlZuzahlung$ = ret$

'ch$ = Left$(tRec!Zuzahlung, 1)
'If (tRec!Warengruppe = 14) Then ch$ = "1"
'If (ch$ = "9") Then
'    MachTaxeAuswahlZuzahlung$ = "20%"
'ElseIf (ch$ = "0") Then
'    MachTaxeAuswahlZuzahlung$ = "kA"
'ElseIf (ch$ = "1") Then
'    MachTaxeAuswahlZuzahlung$ = "nb"
'ElseIf (ch$ = "2") Then
'    MachTaxeAuswahlZuzahlung$ = "ng"
'Else
'    ZuzWert! = 0
'    For i% = 1 To 5
'        zz% = val(Mid$(tRec!Zuzahlung, i%, 1))
'        ZuzWert! = ZuzWert! + Para1.ZuzWert(zz%)
'    Next i%
'
'    AVP! = tRec!VK / 100
'    If ((tRec!FestKz) And (tRec!FESTBETRAG < AVP!)) Then
'        VglWert! = tRec!FESTBETRAG
'    Else
'        VglWert! = AVP!
'    End If
'
'    If ((VglWert! > 0!) And (ZuzWert! > VglWert!)) Then
'        MachTaxeAuswahlZuzahlung$ = "VK"
'    Else
'        MachTaxeAuswahlZuzahlung$ = Format$(ZuzWert!, "#0.00")
'    End If
'End If

Call clsError.DefErrPop
End Function


