VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMerkzettel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim dKurz$
Dim dBezeichnung$
Dim dName$
Dim dZugriff$
Dim dAnzFields%

Dim fKurz$()
Dim fBezeichnung$()


Private Const DefErrModul = "MERKZETTEL.CLS"

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

dKurz$ = "MERK"
dBezeichnung$ = "Merkzettel"
dName$ = "Merkzett.mdb"
dZugriff$ = ""

Set Merkzett1 = Me

Call clsError.DefErrPop
End Sub

Function OpenDatenbank%(DBName$, Optionen%, Schreibgeschuetzt%, DB As Database)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatenbank%")
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
Dim i%, erg%, ErrNumber%

If (DBName$ = "") Then
    DBName$ = dName$
End If

On Error Resume Next
Err.Clear
Set MerkzettelDB = OpenDatabase(DBName$, Optionen%, Schreibgeschuetzt%)
ErrNumber% = Err.Number
On Error GoTo DefErr
If (ErrNumber% = 0) Then
    dAnzFields% = MerkzettelDB.TableDefs("MERKZETTEL").Fields.Count '+ 1
    ReDim fBezeichnung$(dAnzFields% + 1)
'    fBezeichnung$(0) = "ImportKz"
    For i% = 0 To (dAnzFields% - 1)
        fBezeichnung$(i%) = MerkzettelDB.TableDefs("MERKZETTEL").Fields(i%).Name
    Next i%

    Set DB = MerkzettelDB
End If

OpenDatenbank% = ErrNumber%

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
    
'Set fldRecordset = WawiDatRec.Fields(FeldName$)
'Select Case iFeldName$
'    Case "PZN"
'        ret$ = Format(fldRecordset.Value, "0000000")
'    Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT"
'        ret$ = Format(fldRecordset.Value / 100#, "0.00")
'        If (ret$ = "0.00") Or (ret$ = "0,00") Then
'            ret$ = "0"
'        End If
'    Case "ZUZAHLUNG"
'        ret$ = MachTaxeAuswahlZuzahlung$(TaxeRec)
'    Case "GUELTIG", "LPREISAENDERUNG"
'        If (iFeldName$ = "GUELTIG") Then
'            h$ = LTrim$(fldRecordset.Value)
'        Else
'            h$ = Right$(String(6, "0") + Format(fldRecordset.Value, "0"), 6)
'        End If
'        ret$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "."
'        If (Val(Right$(h$, 2)) < 50) Then
'            ret$ = ret$ + "20"
'        Else
'            ret$ = ret$ + "19"
'        End If
'        ret$ = ret$ + Right$(h$, 2)
'    Case Else
'        ret$ = LTrim$(fldRecordset.Value)
'End Select

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
DateiBezeichnung$ = dBezeichnung$
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get DateiName$()
DateiName$ = dName$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property
