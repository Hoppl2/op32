VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAMPV"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 24

Private Type AmpvErstType
    max As Integer
    rest As String * 22
End Type
Private Type AmpvType
  bis As String * 8
  plus As String * 8
  prozent As String * 8
End Type

Dim ampv() As AmpvType
Dim ErstSatz As AmpvErstType

Dim dName$
Dim dZugriff$
Dim dModus$
Dim dSatzLen%
Dim dHandle%

Dim iAnz%

Private Const DefErrModul = "AMPV.CLS"

Private Sub MakeAMPV()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeAMPV")
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
Dim fh%
Dim a As AmpvType, a1 As AmpvErstType

fh% = FreeFile

Open "AMPV.DAT" For Random Access Read Write Shared As #fh% Len = 24
a1.max = 14
a1.rest = String(Len(a1.rest), 0)
Put #fh%, 1, a1

a.bis = clsOpTool.MKD(1.23)
a.prozent = clsOpTool.MKD(68#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 2, a
a.bis = clsOpTool.MKD(1.34)
a.plus = clsOpTool.MKD(0.83)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 3, a
a.bis = clsOpTool.MKD(3.89)
a.prozent = clsOpTool.MKD(62#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 4, a
a.bis = clsOpTool.MKD(4.22)
a.plus = clsOpTool.MKD(2.41)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 5, a
a.bis = clsOpTool.MKD(7.3)
a.prozent = clsOpTool.MKD(57#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 6, a
a.bis = clsOpTool.MKD(8.67)
a.plus = clsOpTool.MKD(4.16)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 7, a
a.bis = clsOpTool.MKD(12.14)
a.prozent = clsOpTool.MKD(48#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 8, a
a.bis = clsOpTool.MKD(13.55)
a.plus = clsOpTool.MKD(5.83)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 9, a
a.bis = clsOpTool.MKD(19.43)
a.prozent = clsOpTool.MKD(43#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 10, a
a.bis = clsOpTool.MKD(22.58)
a.plus = clsOpTool.MKD(8.35)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 11, a
a.bis = clsOpTool.MKD(29.14)
a.prozent = clsOpTool.MKD(37#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 12, a
a.bis = clsOpTool.MKD(35.94)
a.plus = clsOpTool.MKD(10.78)
a.prozent = clsOpTool.MKD(0#)
Put #fh%, 13, a
a.bis = clsOpTool.MKD(543.91)
a.prozent = clsOpTool.MKD(30#)
a.plus = clsOpTool.MKD(0#)
Put #fh%, 14, a
a.plus = clsOpTool.MKD(118.24)
a.prozent = clsOpTool.MKD(8.263)
a.bis = clsOpTool.MKD(0#)
Put #fh%, 15, a
Close #fh%


Call clsError.DefErrPop
End Sub

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
Dim i%

dName$ = "ampv.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN

iAnz% = 0

If (Dir(RTrim$(dName$)) = "") Then Call MakeAMPV

dHandle% = clsDat.FileOpen(RTrim$(dName$), RTrim$(dZugriff$), dModus$, dSatzLen%)
If (dHandle% > 0) Then
    Get #dHandle%, , ErstSatz
    iAnz% = ErstSatz.max
    ReDim ampv(iAnz%)
    For i% = 1 To iAnz%
        Get #dHandle%, , ampv(i%)
    Next i%
    Close #dHandle%
End If

Call clsError.DefErrPop
End Sub

Function CalcAMPV#(BasisWert#, mw%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CalcAMPV#")
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
Dim j%
Dim dpreis#, plus#, prozent#
Dim h$

dpreis# = BasisWert#
plus# = 0#
prozent# = 0

For j% = 1 To UBound(ampv)
    h = h + CStr(clsOpTool.CVD(ampv(j%).bis)) + " " + CStr(clsOpTool.CVD(ampv(j%).plus)) + " " + CStr(clsOpTool.CVD(ampv(j%).prozent)) + vbCrLf
    If (dpreis# <= clsOpTool.CVD(ampv(j%).bis)) Or (clsOpTool.CVD(ampv(j%).bis) = 0#) Then
        plus# = clsOpTool.CVD(ampv(j%).plus)
        prozent# = clsOpTool.CVD(ampv(j%).prozent)
        Exit For
    End If
Next j%

dpreis# = dpreis# * ((100# + prozent#) / 100#) + plus#
dpreis# = dpreis# * ((100# + CDbl(mw%)) / 100#)
CalcAMPV# = Int(dpreis# * 100# + 0.501) / 100#

Call clsError.DefErrPop
End Function


