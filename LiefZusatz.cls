VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiefZusatz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 384

'Private Type LiefFuerHerstStruct
'    lief As Byte
'    herst(2) As String * 5
'End Type
'Dim AlleLiefFuerHerst() As LiefFuerHerstStruct


Private Type RabattTabelle
    Schwellwert As Long
    Rabatt As Integer
End Type
Private Type BmEkTabelle
    AbBM As Integer
    BisBM As Integer
    Vergleich As Byte
    zWert As Double
    Rabatt As Integer
End Type
Private Type LieferantenZusatzStruct
    MindUmsatz As Double
    TabTyp As Byte
    RabTab(4) As RabattTabelle
    BmEkTab(4) As BmEkTabelle
    HatAusnahmen As Byte
    AusnahmenKz(6) As Byte
    AusnahmenWg(9) As Integer
    AusnahmenBM As Integer
    AusnahmenZW As Double
    AusnahmenHerst(5) As String * 5
    
    OrgBeobachtSendungen As Integer
    OrgBeobachtUmsatz(1) As Double
    BeobachtSendungen As Integer
    BeobachtUmsatz(1) As Double
    UmsatzProSendung(1) As Double
    
    SendungenAlt As Integer
    SendungenPlan As Integer
    AliquotProzent As Double
    
    UmsatzBisher(1) As Double
    PrognoseUmsatz(1) As Double
    PrognoseSchwellwert As Double
    PrognoseRabatt As Integer

    WumsatzLief As Byte
    
    DirektMindestBM As Integer
    DirektMindestWert As Long
    KostenProLieferung As Integer
    dummy As String * 6
    DepotPflicht As Byte
    Lieferfrist As Integer
    
    BevorratungsZeitraum As Integer
    ValutaStellung As Integer
    FakturenRabatt As Integer
    FakturenRabattTyp As Byte

    TempBevorratungsZeitraum As Integer
    TempValutaStellung As Integer
    TempFakturenRabatt As Integer
    
    LiefTyp As Byte
    
    MvdaAktiv As Byte
    MvdaProz As Integer
    MvdaLief As Byte
    
    DirektMindestBewertung As Integer
    
    LiefFuerHerst(2) As String * 5
    
    ProzentPlus As Byte
    ZusatzAufwand As Integer
    
    AutoDirekt As Byte
    
    AusnahmenAmRx As Byte
    BmEkTabGruppe(4) As Integer
    
    PreisBasis As Byte
    
    FixRabattAktiv As Byte
    FixRabatt(1) As Integer
    
    rest As String * 22
End Type
Dim lifZusatz As LieferantenZusatzStruct
Dim lifZusatz2 As LieferantenZusatzStruct


Dim iName$(2)
Dim iHandle%(2)
Dim iAnzIndex%

Dim dKurz$
Dim dBezeichnung$
Dim dName$
Dim dZugriff$
Dim dModus$
Dim dSatzLen%
Dim dHandle%
Dim dAnzFields%

Dim fKurz$()
Dim fBezeichnung$()
Dim fStart%()
Dim fLaenge%()
Dim fTyp$()
Dim fKonv%()

Dim FabsErrf%
Dim iFabsRecno&

Dim AktData As String * SATZLEN

Dim iAktRecno&

Dim AuswahlBestLief() As LieferantenZusatzStruct
Dim AnzBestLief%

Dim LIFZUS_INI_DATEI As String

Dim iDirektBestKz As Byte
Dim iBestFax$
Dim iBestModem$
Dim iBestMail$
Dim iAngebotWWW$

Dim iLiefFuerHerst(20) As String * 5
Dim iAnzHerst%
    
    
Dim iMsv3Aktiv As Byte
Dim iVersion20Aktiv As Byte
Dim iAktiveURL As Byte
Dim iEchtURL$
Dim iEchtKundenNr$
Dim iEchtPasswort$
Dim iTestURL$
Dim iTestKundenNr$
Dim iTestPasswort$
Dim iAuftragsartNormal As Byte
Dim iAuftragsartStapel As Byte
Dim iAuftragsartSonder As Byte
Dim iAuftragsartVersand As Byte
Dim iSubstitutionsvereinbarung As Byte
Dim iVerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart As Byte
Dim iVerfuegbarkeitBulkVereinbart As Byte
Dim iRuecknahmeangebotVereinbart As Byte
Dim iGueltigAb$
Dim iVereinbarteAuftragsKennung$(4)
Dim iLieferVorgabe  As Byte
Dim iAuftragsKennung$
Dim iGhCheckBestellen  As Byte
Dim iGhCheckLieferVorgabe As Byte
Dim iGhCheckAuftragKennung$

Private Const DefErrModul = "LIEFZUSATZ.CLS"

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
dKurz$ = "LIEFZUS"
dBezeichnung$ = "Lieferanten-Zusatz"
dName$ = "ldatei2.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(lif)

dAnzFields% = 2
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)


fKurz$(0) = "MINDUMS"
fBezeichnung$(0) = "Mindest-Umsatz"
fStart%(0) = 1
fLaenge%(0) = 8
fTyp$(0) = "N"
fKonv%(0) = 2

fKurz$(1) = "REST"
fBezeichnung$(1) = "Rest"
fStart%(1) = 9
fLaenge%(1) = 120
fTyp$(1) = "T"
fKonv%(1) = 0

iAnzIndex% = 0


Set LifZus1 = Me

LIFZUS_INI_DATEI = CurDir + "\ldatei2.ini"
   
Call clsError.DefErrPop
End Sub

Sub FreeClass()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FreeClass")
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
Set LifZus1 = Nothing
Call clsError.DefErrPop
End Sub

Function OpenDatei%(Optional DateiZugriff$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatei%")
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
Dim iZugriff$
Dim LeerData As String * SATZLEN

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
    If (dHandle% = 0) Then
        OpenDatei% = 1
        Call clsError.DefErrPop: Exit Function
    End If
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
     If (FabsErrf%) Then
        OpenDatei% = 1
        Call clsError.DefErrPop: Exit Function
    End If
Next i%

If (InStr(iZugriff$, "W") > 0) Then
    If (AnzRec < 250) Then
        LeerData = String(SATZLEN, 0)
        For i% = (AnzRec + 1) To 250
            Put #dHandle%, i% + 1, LeerData
        Next i%
    End If
End If

OpenDatei% = 0

Call clsError.DefErrPop
End Function

Sub CloseDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CloseDatei")
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

If (dHandle% > 0) Then
    Close #dHandle%
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexClose(iHandle%(i%))
Next i%

Call clsError.DefErrPop
End Sub

Sub SatzLock(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SatzLock")
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
Call clsDat.iLock(dHandle%, recno&)
Call clsError.DefErrPop
End Sub

Sub SatzUnLock(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SatzUnLock")
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
Call clsDat.iUnLock(dHandle%, recno&)
Call clsError.DefErrPop
End Sub

Sub GetRecord(Optional recno& = 0, Optional EinlesenFlag% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetRecord")
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
Dim temp$

iAktRecno& = recno&
If (LieferantenDBok) And (EinlesenFlag = 0) Then
    temp$ = String(SATZLEN, 0)
    CopyMemory lifZusatz, ByVal temp$, SATZLEN
    Call LieferantenDB1.Konvert2Dos(recno - 1)
Else
    If (recno& > 0) Then
        Seek #dHandle%, recno&
    End If
    Get #dHandle%, , lifZusatz
    
    Call GetDafueNummern(recno& - 1)
End If

Call clsError.DefErrPop
End Sub

Sub PutRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutRecord")
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
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Put #dHandle%, , lifZusatz

Call PutDafueNummern(recno& - 1)

Call clsError.DefErrPop
End Sub

Sub GetInfoRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetInfoRecord")
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
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Get #dHandle%, , AktData
Call clsError.DefErrPop
End Sub

Function IndexSearch%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexSearch%")
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
FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexSearch% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexGeneric%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexGeneric%")
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
FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexGeneric% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexNext%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexNext%")
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
FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexNext% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexPrevious%")
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
FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexPrevious% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexInsert%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexInsert%")
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
FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexInsert% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexDelete%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexDelete%")
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
FabsErrf% = clsDat.IndexDelete%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexDelete% = FabsErrf%
Call clsError.DefErrPop
End Function

Function SucheFeldInd%(Typ%, SuchFeld$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SucheFeldInd%")
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
Dim i%, ret%

ret% = -1

For i% = 0 To (dAnzFields% - 1)
    If (Typ%) Then
        If (fBezeichnung$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    Else
        If (fKurz$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    End If
Next i%

SucheFeldInd% = ret%

Call clsError.DefErrPop
End Function

Function Info$(iFeld$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Info$")
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
Dim iValue As Variant
Dim ind%
Dim ret$, s$

ret$ = ""

s$ = UCase$(iFeld$)
'If (iIndex% >= 0) Then
'    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
'End If


ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    iValue = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
    ret$ = clsDat.DecodFeld(iValue, fTyp$(ind%), fKonv%(ind%))
End If

Info$ = ret$

Call clsError.DefErrPop
End Function

Public Property Get FeldBezeichnung$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(0, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldBezeichnung$ = ret$

End Property

Public Property Get FeldKurz$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(1, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fKurz$(ind%)
End If

FeldKurz = ret$

End Property

Public Property Get DateiBezeichnung$()
DateiBezeichnung$ = dBezeichnung$
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property

Public Property Get DateiLen&()
DateiLen& = LOF(dHandle%)
End Property

Public Property Get AnzRec&()
If (LieferantenDBok) Then
    AnzRec = 260    '9999
Else
    AnzRec& = (DateiLen& / SATZLEN) - 1
    If (AnzRec > 260) Then
        AnzRec = 260
    End If
End If
End Property

Public Property Get Wert(iFeld$, Optional iIndex% = -1) As Variant
Dim ind%
Dim s$
Dim ret As Variant

ret = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
End If

Wert = ret
End Property

Public Property Let Wert(iFeld$, Optional iIndex% = -1, vNewValue As Variant)
Dim ind%
Dim s$, ret$

ret$ = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret$ = clsDat.MakeFeld(vNewValue, fTyp$(ind%), fKonv%(ind%))
    Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)) = Left$(ret$ + Space$(fLaenge%(ind%)), fLaenge%(ind%))
End If

End Property

Function GetPrognoseSchwellwert#(PrognoseUmsatz#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetPrognoseSchwellwert#")
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
GetPrognoseSchwellwert# = GetPrognoseWert(PrognoseUmsatz#, 0)
Call clsError.DefErrPop
End Function

Function GetPrognoseRabatt#(PrognoseUmsatz#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetPrognoseRabatt#")
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
GetPrognoseRabatt# = GetPrognoseWert(PrognoseUmsatz#, 1)
Call clsError.DefErrPop
End Function

Function GetPrognoseWert#(PrognoseUmsatz#, modus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetPrognoseWert#")
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
Dim ret#
                  
ret# = 0#

If (lifZusatz.TabTyp = 0) Then
    For i% = 0 To 4
        With lifZusatz.RabTab(i%)
            If (.Rabatt = 0#) Then Exit For
            If (PrognoseUmsatz# >= .Schwellwert) Then
                If (modus% = 0) Then
                    ret# = .Schwellwert
                Else
                    ret# = .Rabatt / 100#
                End If
            Else
                Exit For
            End If
        End With
    Next i%
    If (modus% = 0) And (ret# = 0#) Then ret# = lifZusatz.RabTab(i%).Schwellwert
End If

GetPrognoseWert# = ret#

Call clsError.DefErrPop
End Function

Sub InitBestLief(pos%, iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitBestLief")
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
Dim temp$, temp2$

AnzBestLief% = pos% + 1
ReDim Preserve AuswahlBestLief(pos%)

If (LieferantenDBok) Then
    temp2$ = String(SATZLEN, 0)
    CopyMemory ByVal temp2$, lifZusatz, SATZLEN
    temp$ = String(SATZLEN, 0)
    CopyMemory lifZusatz, ByVal temp$, SATZLEN
    Call LieferantenDB1.Konvert2Dos(iLief)
    CopyMemory AuswahlBestLief(pos%), lifZusatz, SATZLEN
    CopyMemory lifZusatz, ByVal temp2$, SATZLEN
Else
    Get #dHandle%, iLief% + 1, AuswahlBestLief(pos%)
End If

Call clsError.DefErrPop
End Sub

Function GetBestLief%(sLiefs$, pzn$, art%, lm%, BestRabatt#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetBestLief%")
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
Dim i%, j%, k%, l%, m%, gef%, KriegtRabatt%, ret%, GruppeOk%, val2%
Dim FabsRecno&
Dim aep#, zWert#, RAB#
Dim SQLStr$
                  
'damit im falle 'gar kein Rabatt' ein korrekter Lieferant retourniert wird
ret% = 0
For i% = 0 To (AnzBestLief% - 1)
    If (Mid$(sLiefs$, i% + 1, 1) = "1") Then
        ret% = i%
        Exit For
    End If
Next i%

BestRabatt# = 0#
gef% = False

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
If (TaxeAdoDBok) Then
    On Error Resume Next
    TaxeAdoRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
    If (TaxeAdoRec.EOF = False) Then
        Call clsOpTool.Taxe2ast(pzn$)
        gef% = True
    Else
        If (ArtikelDBok) Then
            SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
            FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
            If (FabsErrf% = 0) Then
                gef% = True
            End If
        Else
            FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Ast1.GetRecord (FabsRecno& + 1)
                gef% = True
            End If
        End If
    End If
Else
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
        Call clsOpTool.Taxe2ast(pzn$)
        gef% = True
    Else
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Ast1.GetRecord (FabsRecno& + 1)
            gef% = True
        End If
    End If
End If

If (gef%) Then aep# = Ast1.aep


'For i% = 0 To (AnzBestLief% - 1)
'    If (Mid$(sLiefs$, i% + 1, 1) = "1") And (AuswahlBestLief(i%).HatAusnahmen) Then
'        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'        If (TaxeRec.EOF = False) Then
'    '        AEP# = TaxeRec!EK / 100#
'            Call clsOpTool.Taxe2ast(pzn$)
'            gef% = True
'        Else
'            FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
'            If (FabsErrf% = 0) Then
'                Ast1.GetRecord (FabsRecno& + 1)
'    '            AEP# = Ast1.AEP
'                gef% = True
'            End If
'        End If
'
'        If (gef%) Then aep# = Ast1.aep
'        Exit For
'    End If
'Next i%

If (gef%) Then
    For i% = 0 To (AnzBestLief% - 1)
        If (Mid$(sLiefs$, i% + 1, 1) = "1") Then
            With AuswahlBestLief(i%)
                KriegtRabatt% = True
                If (.HatAusnahmen) Then
                    KriegtRabatt% = CheckAusnahme%(i%, art%, lm%, aep#)
                End If
                If (KriegtRabatt%) Then
                    If (.TabTyp = 0) Then
                        RAB# = .PrognoseRabatt
                    Else
                        For j% = 0 To 4
                            With AuswahlBestLief(i%).BmEkTab(j%)
                                If (.Rabatt = 0#) Then Exit For
                                If (lm% >= .AbBM) And ((lm% <= .BisBM) Or (.BisBM = 0)) Then
                                    zWert# = aep# * lm%
                                    If ((.Vergleich = 0) And (zWert# < .zWert)) Or _
                                       ((.Vergleich = 1) And (zWert# <= .zWert)) Or _
                                       ((.Vergleich = 2) And (zWert# = .zWert)) Or _
                                       ((.Vergleich = 3) And (zWert# >= .zWert)) Or _
                                       ((.Vergleich > 3) And (zWert# > .zWert)) Then
                                        GruppeOk% = True
                                        val2% = AuswahlBestLief(i%).BmEkTabGruppe(j%)
                                        If (TaxeAdoDBok) Then
                                            If (TaxeAdoRec.EOF = False) And (val2% > 0) Then
                                                GruppeOk% = 0
                                                l% = Format(TaxeAdoRec!AbgabeBest, "0")
                                                k% = 1
                                                For m% = 1 To l%
                                                    k% = k% * 2
                                                Next m%
                                                If (val2% And k%) Then
                                                    GruppeOk% = True
                                                End If
                                            End If
                                        Else
                                            If (TaxeRec.EOF = False) And (val2% > 0) Then
                                                GruppeOk% = 0
                                                l% = Format(TaxeRec!AbgabeBest, "0")
                                                k% = 1
                                                For m% = 1 To l%
                                                    k% = k% * 2
                                                Next m%
                                                If (val2% And k%) Then
                                                    GruppeOk% = True
                                                End If
                                            End If
                                        End If
                                        If (GruppeOk%) Then
                                            RAB# = .Rabatt
                                            Exit For
                                        End If
                                    End If
                                End If
                            End With
                        Next j%
                    End If
                    If (RAB# > BestRabatt#) Then
                        BestRabatt# = RAB#
                        ret% = i%
                    End If
                End If
            End With
        End If
    Next i%
End If
   
BestRabatt# = BestRabatt# / 100!
GetBestLief% = ret%

Call clsError.DefErrPop
End Function

Function CheckAusnahme%(pos%, art%, lm%, aep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckAusnahme%")
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
Dim ret%

If (pos% < 0) Then
    ret% = CheckAusnahme1%(lifZusatz, art%, lm%, aep#)
Else
    ret% = CheckAusnahme1%(AuswahlBestLief(pos), art%, lm%, aep#)
End If

CheckAusnahme% = ret%

Call clsError.DefErrPop
End Function

Private Function CheckAusnahme1%(lz As LieferantenZusatzStruct, art%, lm%, aep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckAusnahme1%")
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
Dim i%, KriegtRabatt%, wg%
Dim herst$

With lz

    KriegtRabatt% = True

    If (.AusnahmenKz(0) And (Ast1.lac = "1" Or Ast1.lac = "4")) Then KriegtRabatt% = 0
    If (.AusnahmenKz(1) And (art% = 2)) Then KriegtRabatt% = 0
    If (KriegtRabatt% And .AusnahmenKz(2) And InStr(Ast1.rez, "SG") <> 0) Then KriegtRabatt% = 0
    If (KriegtRabatt% And .AusnahmenKz(3)) Then
        For i% = 0 To 9
            If (.AusnahmenWg(i%) <> 0) Then
                wg% = Val(Ast1.wg1)
                If (.AusnahmenWg(i%) > 9) Then wg% = wg% * 10 + Val(Ast1.wg2)
                If (.AusnahmenWg(i%) = wg%) Then
                    KriegtRabatt% = 0
                    Exit For
                End If
            End If
        Next i%
    End If
    If (KriegtRabatt% And .AusnahmenKz(4) And (lm% <= .AusnahmenBM)) Then KriegtRabatt% = 0
    If (KriegtRabatt% And .AusnahmenKz(5) And (aep# * lm% <= .AusnahmenZW)) Then KriegtRabatt% = 0
    If (KriegtRabatt% And .AusnahmenKz(6)) Then
        herst$ = Ast1.herst: If herst$ = Space$(Len(herst$)) Then herst$ = ""
        For i% = 0 To 5
            If (herst$ = .AusnahmenHerst(i%)) Then
                KriegtRabatt% = 0
                i% = 5
            End If
        Next i%
    End If
    If (KriegtRabatt% And (.AusnahmenAmRx = 101)) Then
        If (InStr(Ast1.rez, "+") <> 0) Then
            KriegtRabatt% = 0
        End If
    End If
End With

CheckAusnahme1% = KriegtRabatt%

Call clsError.DefErrPop
End Function

Function GlobalPrognoseRabatt#(art%, lm%, aep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GlobalPrognoseRabatt#")
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
Dim i%, j%, k%, l%, GruppeOk%, val2%
Dim ret#, GesAEP#

ret# = 0#

If (CheckAusnahme1%(lifZusatz, art%, lm%, aep#)) Then
    With LifZus1
        If (.TabTyp = 0) Then
            ret# = .PrognoseRabatt
        Else
            For i% = 0 To 4
                If (.BmEkRabatt(i%) = 0#) Then Exit For
                If (lm% >= .AbBM(i%)) And ((lm% <= .BisBM(i%)) Or (.BisBM(i%) = 0)) Then
                    GesAEP# = aep# * lm%
                    If ((.Vergleich(i%) = 0) And (GesAEP# < .zWert(i%))) Or _
                       ((.Vergleich(i%) = 1) And (GesAEP# <= .zWert(i%))) Or _
                       ((.Vergleich(i%) = 2) And (GesAEP# = .zWert(i%))) Or _
                       ((.Vergleich(i%) = 3) And (GesAEP# >= .zWert(i%))) Or _
                       ((.Vergleich(i%) > 3) And (GesAEP# > .zWert(i%))) Then
                        GruppeOk% = True
                        val2% = .BmEkTabGruppe(i%)
                        If (TaxeAdoDBok) Then
                            If (TaxeAdoRec.EOF = False) And (val2% > 0) Then
                                GruppeOk% = 0
                                l% = Format(TaxeAdoRec!AbgabeBest, "0")
                                k% = 1
                                For j% = 1 To l%
                                    k% = k% * 2
                                Next j%
                                If (val2% And k%) Then
                                    GruppeOk% = True
                                End If
                            End If
                        Else
                            If (TaxeRec.EOF = False) And (val2% > 0) Then
                                GruppeOk% = 0
                                l% = Format(TaxeRec!AbgabeBest, "0")
                                k% = 1
                                For j% = 1 To l%
                                    k% = k% * 2
                                Next j%
                                If (val2% And k%) Then
                                    GruppeOk% = True
                                End If
                            End If
                        End If
                        If (GruppeOk%) Then
                            ret# = .BmEkRabatt(i%)
                            Exit For
                        End If
                    End If
                End If
            Next i%
        End If
    End With
End If

GlobalPrognoseRabatt# = ret#

Call clsError.DefErrPop
End Function

Sub InitDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitDatei")
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
Dim LeerData As String * SATZLEN
    
LeerData = String(SATZLEN, 0)
For i% = 0 To 250
    Put #dHandle%, , LeerData
Next i%

Call clsError.DefErrPop
End Sub

Function GetWumsatzLief%(OrgLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetWumsatzLief%")
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
Dim ret%, iWumsatzLief%, iLief%, iDatenVonLief%
Dim SQLStr$
Dim LieferantenRec2 As New ADODB.Recordset
Dim LieferantenzusatzRec2 As New ADODB.Recordset
    
ret% = OrgLief%
If (LieferantenDBok) Then
    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(ret)
    LieferantenRec2.Open SQLStr, LieferantenDB1.ActiveConn  ' LieferantenConn
    If (LieferantenRec2.RecordCount <> 0) Then
        iDatenVonLief = clsOpTool.CheckNullLong(LieferantenRec2!DatenVonLief)
        If (iDatenVonLief > 0) Then
            ret = iDatenVonLief
        End If
    End If
    LieferantenRec2.Close
    
    SQLStr$ = "SELECT * FROM LieferantenZusatz WHERE LiefNr =" + Str$(ret)
    LieferantenzusatzRec2.Open SQLStr, LieferantenDB1.ActiveConn ' LieferantenConn
    If Not (LieferantenzusatzRec2.EOF) Then
        iWumsatzLief = clsOpTool.CheckNullLong(LieferantenzusatzRec2!WumsatzLief)
        If (iWumsatzLief > 0) Then  'And (iWumsatzLief <> OrgLief%) Then
            ret% = iWumsatzLief
        End If
    End If
    LieferantenzusatzRec2.Close
Else
    Get #dHandle%, OrgLief% + 1, lifZusatz2
    If (lifZusatz2.WumsatzLief > 0) And (lifZusatz2.WumsatzLief <> OrgLief%) Then
        ret% = lifZusatz2.WumsatzLief
    End If
End If

GetWumsatzLief% = ret%

Call clsError.DefErrPop
End Function

Public Property Get FabsRecno&()
FabsRecno& = iFabsRecno&
End Property



Public Property Get MindestUmsatz#()
MindestUmsatz# = lifZusatz.MindUmsatz
End Property
Public Property Let MindestUmsatz(ByVal vNewValue#)
lifZusatz.MindUmsatz = vNewValue#
End Property

Public Property Get TabTyp() As Byte
TabTyp = lifZusatz.TabTyp
End Property
Public Property Let TabTyp(ByVal vNewValue As Byte)
lifZusatz.TabTyp = vNewValue
End Property

Public Property Get Schwellwert&(ByVal ind%)
Schwellwert& = lifZusatz.RabTab(ind%).Schwellwert
End Property
Public Property Let Schwellwert(ByVal ind%, ByVal vNewValue&)
lifZusatz.RabTab(ind%).Schwellwert = vNewValue&
End Property

Public Property Get Rabatt#(ByVal ind%)
Rabatt# = lifZusatz.RabTab(ind%).Rabatt / 100
End Property
Public Property Let Rabatt(ByVal ind%, ByVal vNewValue#)
lifZusatz.RabTab(ind%).Rabatt = CInt(vNewValue# * 100)
End Property


Public Property Get AbBM%(ByVal ind%)
AbBM% = lifZusatz.BmEkTab(ind%).AbBM
End Property
Public Property Let AbBM(ByVal ind%, ByVal vNewValue%)
lifZusatz.BmEkTab(ind%).AbBM = vNewValue%
End Property

Public Property Get BisBM%(ByVal ind%)
BisBM% = lifZusatz.BmEkTab(ind%).BisBM
End Property
Public Property Let BisBM(ByVal ind%, ByVal vNewValue%)
lifZusatz.BmEkTab(ind%).BisBM = vNewValue%
End Property

Public Property Get Vergleich(ByVal ind%) As Byte
Vergleich = lifZusatz.BmEkTab(ind%).Vergleich
End Property
Public Property Let Vergleich(ByVal ind%, ByVal vNewValue As Byte)
lifZusatz.BmEkTab(ind%).Vergleich = vNewValue
End Property

Public Property Get zWert#(ByVal ind%)
zWert# = lifZusatz.BmEkTab(ind%).zWert
End Property
Public Property Let zWert(ByVal ind%, ByVal vNewValue#)
lifZusatz.BmEkTab(ind%).zWert = vNewValue#
End Property

Public Property Get BmEkRabatt#(ByVal ind%)
BmEkRabatt# = lifZusatz.BmEkTab(ind%).Rabatt / 100
End Property
Public Property Let BmEkRabatt(ByVal ind%, ByVal vNewValue#)
lifZusatz.BmEkTab(ind%).Rabatt = CInt(vNewValue# * 100)
End Property


Public Property Get HatAusnahmen() As Byte
HatAusnahmen = lifZusatz.HatAusnahmen
End Property
Public Property Let HatAusnahmen(ByVal vNewValue As Byte)
lifZusatz.HatAusnahmen = vNewValue
End Property

Public Property Get AusnahmenKz(ByVal ind%) As Byte
AusnahmenKz = lifZusatz.AusnahmenKz(ind%)
End Property
Public Property Let AusnahmenKz(ByVal ind%, ByVal vNewValue As Byte)
lifZusatz.AusnahmenKz(ind%) = vNewValue
End Property

Public Property Get AusnahmenWg%(ByVal ind%)
AusnahmenWg% = lifZusatz.AusnahmenWg(ind%)
End Property
Public Property Let AusnahmenWg(ByVal ind%, ByVal vNewValue%)
lifZusatz.AusnahmenWg(ind%) = vNewValue%
End Property

Public Property Get AusnahmenBM%()
AusnahmenBM% = lifZusatz.AusnahmenBM
End Property
Public Property Let AusnahmenBM(ByVal vNewValue%)
lifZusatz.AusnahmenBM = vNewValue%
End Property

Public Property Get AusnahmenZW#()
AusnahmenZW# = lifZusatz.AusnahmenZW
End Property
Public Property Let AusnahmenZW(ByVal vNewValue#)
lifZusatz.AusnahmenZW = vNewValue#
End Property

Public Property Get AusnahmenHerst$(ByVal ind%)
Dim i%, h$
h$ = lifZusatz.AusnahmenHerst(ind%)
For i% = 1 To (Len(h$))
    If (Mid$(h$, i%, 1) = Chr$(0)) Then Mid$(h$, i%, 1) = " "
Next i%
AusnahmenHerst$ = h$
End Property
Public Property Let AusnahmenHerst(ByVal ind%, ByVal vNewValue$)
lifZusatz.AusnahmenHerst(ind%) = Left$(Trim$(vNewValue$) + Space$(5), 5)
End Property


Public Property Get OrgBeobachtSendungen%()
OrgBeobachtSendungen% = lifZusatz.OrgBeobachtSendungen
End Property
Public Property Let OrgBeobachtSendungen(ByVal vNewValue%)
lifZusatz.OrgBeobachtSendungen = vNewValue%
End Property

Public Property Get OrgBeobachtUmsatz#(ByVal ind%)
OrgBeobachtUmsatz# = lifZusatz.OrgBeobachtUmsatz(ind%)
End Property
Public Property Let OrgBeobachtUmsatz(ByVal ind%, ByVal vNewValue#)
lifZusatz.OrgBeobachtUmsatz(ind%) = vNewValue#
End Property

Public Property Get BeobachtSendungen%()
BeobachtSendungen% = lifZusatz.BeobachtSendungen
End Property
Public Property Let BeobachtSendungen(ByVal vNewValue%)
lifZusatz.BeobachtSendungen = vNewValue%
End Property

Public Property Get BeobachtUmsatz#(ByVal ind%)
BeobachtUmsatz# = lifZusatz.BeobachtUmsatz(ind%)
End Property
Public Property Let BeobachtUmsatz(ByVal ind%, ByVal vNewValue#)
lifZusatz.BeobachtUmsatz(ind%) = vNewValue#
End Property

Public Property Get UmsatzProSendung#(ByVal ind%)
UmsatzProSendung# = lifZusatz.UmsatzProSendung(ind%)
End Property
Public Property Let UmsatzProSendung(ByVal ind%, ByVal vNewValue#)
lifZusatz.UmsatzProSendung(ind%) = vNewValue#
End Property


Public Property Get SendungenAlt%()
SendungenAlt% = lifZusatz.SendungenAlt
End Property
Public Property Let SendungenAlt(ByVal vNewValue%)
lifZusatz.SendungenAlt = vNewValue%
End Property

Public Property Get SendungenPlan%()
SendungenPlan% = lifZusatz.SendungenPlan
End Property
Public Property Let SendungenPlan(ByVal vNewValue%)
lifZusatz.SendungenPlan = vNewValue%
End Property

Public Property Get AliquotProzent#()
AliquotProzent# = lifZusatz.AliquotProzent
End Property
Public Property Let AliquotProzent(ByVal vNewValue#)
lifZusatz.AliquotProzent = vNewValue#
End Property


Public Property Get UmsatzBisher#(ByVal ind%)
UmsatzBisher# = lifZusatz.UmsatzBisher(ind%)
End Property
Public Property Let UmsatzBisher(ByVal ind%, ByVal vNewValue#)
lifZusatz.UmsatzBisher(ind%) = vNewValue#
End Property

Public Property Get PrognoseUmsatz#(ByVal ind%)
PrognoseUmsatz# = lifZusatz.PrognoseUmsatz(ind%)
End Property
Public Property Let PrognoseUmsatz(ByVal ind%, ByVal vNewValue#)
lifZusatz.PrognoseUmsatz(ind%) = vNewValue#
End Property

Public Property Get PrognoseSchwellwert#()
PrognoseSchwellwert# = lifZusatz.PrognoseSchwellwert
End Property
Public Property Let PrognoseSchwellwert(ByVal vNewValue#)
lifZusatz.PrognoseSchwellwert = vNewValue#
End Property

Public Property Get PrognoseRabatt#()
PrognoseRabatt# = lifZusatz.PrognoseRabatt / 100
End Property
Public Property Let PrognoseRabatt(ByVal vNewValue#)
lifZusatz.PrognoseRabatt = CInt(vNewValue# * 100)
End Property


Public Property Get WumsatzLief() As Byte
WumsatzLief = lifZusatz.WumsatzLief
End Property
Public Property Let WumsatzLief(ByVal vNewValue As Byte)
lifZusatz.WumsatzLief = vNewValue
End Property

Public Property Get DirektMindestBM%()
DirektMindestBM% = lifZusatz.DirektMindestBM
End Property
Public Property Let DirektMindestBM(ByVal vNewValue%)
lifZusatz.DirektMindestBM = vNewValue%
End Property

Public Property Get DirektMindestWert&()
DirektMindestWert& = lifZusatz.DirektMindestWert
End Property
Public Property Let DirektMindestWert(ByVal vNewValue&)
lifZusatz.DirektMindestWert = vNewValue&
End Property

Public Property Get KostenProLieferung#()
KostenProLieferung# = lifZusatz.KostenProLieferung / 100
End Property
Public Property Let KostenProLieferung(ByVal vNewValue#)
lifZusatz.KostenProLieferung = CInt(vNewValue# * 100)
End Property

Public Property Get LiefFuerHerst$(ByVal ind%)
If (ind% <= UBound(iLiefFuerHerst)) Then
    LiefFuerHerst$ = iLiefFuerHerst(ind%)
Else
    LiefFuerHerst$ = ""
End If
'LiefFuerHerst$ = lifZusatz.LiefFuerHerst(ind%)
End Property
Public Property Let LiefFuerHerst(ByVal ind%, ByVal vNewValue$)
If (ind% <= UBound(iLiefFuerHerst)) Then
    iLiefFuerHerst(ind%) = vNewValue$
End If
'lifZusatz.LiefFuerHerst(ind%) = vNewValue$
End Property

Public Property Get DepotPflicht() As Byte
DepotPflicht = lifZusatz.DepotPflicht
End Property
Public Property Let DepotPflicht(ByVal vNewValue As Byte)
lifZusatz.DepotPflicht = vNewValue
End Property

Public Property Get Lieferfrist%()
Lieferfrist% = lifZusatz.Lieferfrist
End Property
Public Property Let Lieferfrist(ByVal vNewValue%)
lifZusatz.Lieferfrist = vNewValue%
End Property

Public Property Get BevorratungsZeitraum%()
BevorratungsZeitraum% = lifZusatz.BevorratungsZeitraum
End Property
Public Property Let BevorratungsZeitraum(ByVal vNewValue%)
lifZusatz.BevorratungsZeitraum = vNewValue%
End Property

Public Property Get ValutaStellung%()
ValutaStellung% = lifZusatz.ValutaStellung
End Property
Public Property Let ValutaStellung(ByVal vNewValue%)
lifZusatz.ValutaStellung = vNewValue%
End Property

Public Property Get FakturenRabatt#()
FakturenRabatt# = lifZusatz.FakturenRabatt / 100
End Property
Public Property Let FakturenRabatt(ByVal vNewValue#)
lifZusatz.FakturenRabatt = CInt(vNewValue# * 100)
End Property

Public Property Get FakturenRabattTyp() As Byte
FakturenRabattTyp = lifZusatz.FakturenRabattTyp
End Property
Public Property Let FakturenRabattTyp(ByVal vNewValue As Byte)
lifZusatz.FakturenRabattTyp = vNewValue
End Property

Public Property Get DirektBestFaxKz() As Byte
Dim ret As Byte

ret = 0
'If (lifZusatz.DirektBestKz And &H1) Then ret = 1
If (iDirektBestKz And &H1) Then ret = 1

DirektBestFaxKz = ret
End Property
Public Property Let DirektBestFaxKz(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = lifZusatz.DirektBestKz
iByte = iDirektBestKz
If (vNewValue) Then
    iByte = iByte Or &H1
Else
    iByte = iByte And (&H1 Xor &HFF)
End If
'lifZusatz.DirektBestKz = iByte
iDirektBestKz = iByte
End Property
        
Public Property Get DirektBestModemKz() As Byte
Dim ret As Byte

ret = 0
'If (lifZusatz.DirektBestKz And &H2) Then ret = 1
If (iDirektBestKz And &H2) Then ret = 1

DirektBestModemKz = ret
End Property
Public Property Let DirektBestModemKz(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = lifZusatz.DirektBestKz
iByte = iDirektBestKz
If (vNewValue) Then
    iByte = iByte Or &H2
Else
    iByte = iByte And (&H2 Xor &HFF)
End If
'lifZusatz.DirektBestKz = iByte
iDirektBestKz = iByte
End Property
        
Public Property Get DirektBestMailKz() As Byte
Dim ret As Byte

ret = 0
'If (lifZusatz.DirektBestKz And &H4) Then ret = 1
If (iDirektBestKz And &H4) Then ret = 1

DirektBestMailKz = ret
End Property
Public Property Let DirektBestMailKz(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = lifZusatz.DirektBestKz
iByte = iDirektBestKz
If (vNewValue) Then
    iByte = iByte Or &H4
Else
    iByte = iByte And (&H4 Xor &HFF)
End If
'lifZusatz.DirektBestKz = iByte
iDirektBestKz = iByte
End Property
        
Public Property Get DirektBestDruckKz() As Byte
Dim ret As Byte

ret = 0
'If (lifZusatz.DirektBestKz And &H8) Then ret = 1
If (iDirektBestKz And &H8) Then ret = 1

DirektBestDruckKz = ret
End Property
Public Property Let DirektBestDruckKz(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = lifZusatz.DirektBestKz
iByte = iDirektBestKz
If (vNewValue) Then
    iByte = iByte Or &H8
Else
    iByte = iByte And (&H8 Xor &HFF)
End If
'lifZusatz.DirektBestKz = iByte
iDirektBestKz = iByte
End Property
        
Public Property Get DirektBestComputerFaxKz() As Byte
Dim ret As Byte

ret = 0
'If (lifZusatz.DirektBestKz And &H8) Then ret = 1
If (iDirektBestKz And &H10) Then ret = 1

DirektBestComputerFaxKz = ret
End Property
Public Property Let DirektBestComputerFaxKz(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = lifZusatz.DirektBestKz
iByte = iDirektBestKz
If (vNewValue) Then
    iByte = iByte Or &H10
Else
    iByte = iByte And (&H10 Xor &HFF)
End If
'lifZusatz.DirektBestKz = iByte
iDirektBestKz = iByte
End Property
        
Public Property Get DirektBestFax$()
'DirektBestFax$ = lifZusatz.BestFax
DirektBestFax$ = iBestFax$
End Property
Public Property Let DirektBestFax(ByVal vNewValue$)
'lifZusatz.BestFax = vNewValue$
iBestFax$ = vNewValue$
End Property

Public Property Get DirektBestModem$()
'DirektBestModem$ = lifZusatz.BestModem
DirektBestModem$ = iBestModem$
End Property
Public Property Let DirektBestModem(ByVal vNewValue$)
'lifZusatz.BestModem = vNewValue$
iBestModem$ = vNewValue$
End Property

Public Property Get DirektBestMail$()
'DirektBestMail$ = lifZusatz.BestMail
DirektBestMail$ = iBestMail$
End Property
Public Property Let DirektBestMail(ByVal vNewValue$)
'lifZusatz.BestMail = vNewValue$
iBestMail$ = vNewValue$
End Property

Public Property Get AngebotWWW$()
'AngebotWWW$ = lifZusatz.AngebotWWW
AngebotWWW$ = iAngebotWWW$
End Property
Public Property Let AngebotWWW(ByVal vNewValue$)
'lifZusatz.AngebotWWW = vNewValue$
iAngebotWWW$ = vNewValue$
End Property


Public Property Get TempBevorratungsZeitraum%()
TempBevorratungsZeitraum% = lifZusatz.TempBevorratungsZeitraum
End Property
Public Property Let TempBevorratungsZeitraum(ByVal vNewValue%)
lifZusatz.TempBevorratungsZeitraum = vNewValue%
End Property

Public Property Get TempValutaStellung%()
TempValutaStellung% = lifZusatz.TempValutaStellung
End Property
Public Property Let TempValutaStellung(ByVal vNewValue%)
lifZusatz.TempValutaStellung = vNewValue%
End Property

Public Property Get TempFakturenRabatt#()
TempFakturenRabatt# = lifZusatz.TempFakturenRabatt / 100
End Property
Public Property Let TempFakturenRabatt(ByVal vNewValue#)
lifZusatz.TempFakturenRabatt = CInt(vNewValue# * 100)
End Property

Public Property Get MvdaProzent#()
MvdaProzent# = lifZusatz.MvdaProz / 100#
End Property
Public Property Let MvdaProzent(ByVal vNewValue#)
lifZusatz.MvdaProz = Int(vNewValue# * 100 + 0.5)
End Property

Public Property Get MvdaLief() As Byte
MvdaLief = lifZusatz.MvdaLief
End Property
Public Property Let MvdaLief(ByVal vNewValue As Byte)
lifZusatz.MvdaLief = vNewValue
End Property

Public Property Get DirektMindestBewertung%()
DirektMindestBewertung% = lifZusatz.DirektMindestBewertung
End Property
Public Property Let DirektMindestBewertung(ByVal vNewValue%)
lifZusatz.DirektMindestBewertung = vNewValue%
End Property


Public Property Get IstDirektLieferant() As Byte
Dim ret As Byte, iLiefTyp As Byte
Dim s%, von%, bis%, iLief%, gefunden%
Dim lief$, Such$

'ret = 0
'If (lifZusatz.LiefTyp = 0) Then ret = 1

iLief% = iAktRecno& - 1
iLiefTyp = lifZusatz.LiefTyp
If (iLiefTyp = 100) Or (InStr(Para1.Benutz, "w") = 0) Then
    ret = 0
ElseIf (iLiefTyp = 101) Then
    ret = 1
'ElseIf (Para1.DirektLieferanten <> "") Then
'    ret = 0
'    'Suche, ob der WÜ-Lieferant unter den in PDATEI angeführten Direktlieferanten ist
'    gefunden% = False
'    Lief$ = Para1.DirektLieferanten
'    While ((Lief$ <> "") And Not (gefunden%))
'        s% = InStr(Lief$, ",")
'        If (s% > 0) Then
'            Such$ = Left$(Lief$, s% - 1)
'            Lief$ = Mid$(Lief$, s% + 1)
'        Else
'            Such$ = Lief$
'            Lief$ = ""
'        End If
'
'        s% = InStr(Such$, "-")
'        If (s% > 0) Then
'            von% = val(Left$(Such$, s% - 1))
'            bis% = val(Mid$(Such$, s% + 1))
'            If (bis% = 0) Then bis% = 200
'            If ((von% <= iLief%) And (iLief% <= bis%)) Then gefunden% = True
'        Else
'            If (val(Such$) = iLief%) Then gefunden% = True
'        End If
'    Wend
'    If (gefunden%) Then ret = 1
'ElseIf (iLief% >= 10) Then
'    ret = 1
Else
    ret = 0
End If

IstDirektLieferant = ret
End Property
Public Property Let IstDirektLieferant(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = 1
'If (vNewValue) Then iByte = 0

iByte = 100
If (vNewValue) Then iByte = 101
lifZusatz.LiefTyp = iByte
End Property


Public Property Get IstMvdaLieferant() As Byte
Dim ret As Byte

ret = 0
If (lifZusatz.MvdaAktiv = 101) Then ret = 1

IstMvdaLieferant = ret
End Property
Public Property Let IstMvdaLieferant(ByVal vNewValue As Byte)
Dim iByte As Byte

iByte = 100
If (vNewValue) Then iByte = 101
lifZusatz.MvdaAktiv = iByte
End Property

Public Property Get ProzentPlus() As Byte
ProzentPlus = lifZusatz.ProzentPlus
End Property
Public Property Let ProzentPlus(ByVal vNewValue As Byte)
lifZusatz.ProzentPlus = vNewValue
End Property

Public Property Get ZusatzAufwand#()
ZusatzAufwand# = lifZusatz.ZusatzAufwand / 10
End Property
Public Property Let ZusatzAufwand(ByVal vNewValue#)
lifZusatz.ZusatzAufwand = Int(vNewValue * 10 + 0.501)
End Property

Public Property Get IstAutoDirektLieferant() As Byte
Dim ret As Byte

ret = 0
If (InStr(Para1.Benutz, "w") = 0) Then
ElseIf (lifZusatz.AutoDirekt = 112) Then
    ret = 1
End If

IstAutoDirektLieferant = ret
End Property
'Public Property Let IstAutoDirektLieferant(ByVal vNewValue As Byte)
'Dim iByte As Byte
'
'iByte = 100
'If (vNewValue) Then iByte = 101
'lifZusatz.AutoDirekt = iByte
'End Property

Public Property Get ZuordnungenModus() As Byte
Dim ret As Byte

ret = 1
If (lifZusatz.AutoDirekt >= 110) Then ret = lifZusatz.AutoDirekt - 110

ZuordnungenModus = ret
End Property
Public Property Let ZuordnungenModus(ByVal vNewValue As Byte)
lifZusatz.AutoDirekt = vNewValue + 110
End Property

Public Property Get ZuordnungenAktiv() As Byte
Dim ret As Byte

ret = 0
If (lifZusatz.AutoDirekt >= 111) Then ret = 1

ZuordnungenAktiv = ret
End Property

Public Property Get AusnahmenAmRx() As Byte
Dim ret As Byte

ret = 0
If (lifZusatz.AusnahmenAmRx = 101) Then ret = 1

AusnahmenAmRx = ret
End Property
Public Property Let AusnahmenAmRx(ByVal vNewValue As Byte)
Dim iByte As Byte

iByte = 100
If (vNewValue) Then iByte = 101
lifZusatz.AusnahmenAmRx = iByte
End Property

Public Property Get BmEkTabGruppe%(ByVal ind%)
BmEkTabGruppe% = lifZusatz.BmEkTabGruppe%(ind%)
End Property
Public Property Let BmEkTabGruppe(ByVal ind%, ByVal vNewValue%)
lifZusatz.BmEkTabGruppe(ind%) = vNewValue%
End Property

Public Property Get PreisBasis() As Byte
Dim ret As Byte

ret = 0
If (lifZusatz.PreisBasis = 111) Then
    ret = 1
ElseIf (lifZusatz.PreisBasis = 112) Then
    ret = 2
End If

PreisBasis = ret
End Property
Public Property Let PreisBasis(ByVal vNewValue As Byte)
Dim iByte As Byte

'iByte = 110
'If (vNewValue) Then iByte = 111
iByte = 110 + vNewValue
lifZusatz.PreisBasis = iByte
End Property

Public Property Get HatFixRabatt() As Byte
Dim ret As Byte

ret = 0
If (lifZusatz.FixRabattAktiv = 101) Then ret = 1

HatFixRabatt = ret
End Property
Public Property Let HatFixRabatt(ByVal vNewValue As Byte)
Dim iByte As Byte

iByte = 100
If (vNewValue) Then iByte = 101
lifZusatz.FixRabattAktiv = iByte
End Property

Public Property Get FixRabatt#(ByVal ind%)
FixRabatt# = lifZusatz.FixRabatt(ind) / 100#
End Property
Public Property Let FixRabatt(ByVal ind%, ByVal vNewValue#)
lifZusatz.FixRabatt(ind) = Int(vNewValue# * 100 + 0.5)
End Property

Public Property Get rest$()
rest$ = ""  ' lifZusatz.rest
End Property
Public Property Let rest(ByVal vNewValue$)
'lifZusatz.rest = vNewValue$
End Property

Public Property Get Msv3Aktiv() As Byte
Msv3Aktiv = iMsv3Aktiv
End Property
Public Property Let Msv3Aktiv(ByVal vNewValue As Byte)
iMsv3Aktiv = vNewValue
End Property

Public Property Get Version20Aktiv() As Byte
Version20Aktiv = iVersion20Aktiv
End Property
Public Property Let Version20Aktiv(ByVal vNewValue As Byte)
iVersion20Aktiv = vNewValue
End Property

Public Property Get AktiveURL() As Byte
AktiveURL = iAktiveURL
End Property
Public Property Let AktiveURL(ByVal vNewValue As Byte)
iAktiveURL = vNewValue
End Property

Public Property Get EchtURL$()
EchtURL = iEchtURL
End Property
Public Property Let EchtURL(ByVal vNewValue$)
iEchtURL = vNewValue$
End Property

Public Property Get EchtKundenNr$()
EchtKundenNr = iEchtKundenNr
End Property
Public Property Let EchtKundenNr(ByVal vNewValue$)
iEchtKundenNr = vNewValue$
End Property

Public Property Get EchtPasswort$()
EchtPasswort = iEchtPasswort
End Property
Public Property Let EchtPasswort(ByVal vNewValue$)
iEchtPasswort = vNewValue$
End Property

Public Property Get TestURL$()
TestURL = iTestURL
End Property
Public Property Let TestURL(ByVal vNewValue$)
iTestURL = vNewValue$
End Property

Public Property Get TestKundenNr$()
TestKundenNr = iTestKundenNr
End Property
Public Property Let TestKundenNr(ByVal vNewValue$)
iTestKundenNr = vNewValue$
End Property

Public Property Get TestPasswort$()
TestPasswort = iTestPasswort
End Property
Public Property Let TestPasswort(ByVal vNewValue$)
iTestPasswort = vNewValue$
End Property
    
Public Property Get AuftragsartNormal() As Byte
AuftragsartNormal = iAuftragsartNormal
End Property
Public Property Let AuftragsartNormal(ByVal vNewValue As Byte)
iAuftragsartNormal = vNewValue
End Property

Public Property Get AuftragsartStapel() As Byte
AuftragsartStapel = iAuftragsartStapel
End Property
Public Property Let AuftragsartStapel(ByVal vNewValue As Byte)
iAuftragsartStapel = vNewValue
End Property
    
Public Property Get AuftragsartSonder() As Byte
AuftragsartSonder = iAuftragsartSonder
End Property
Public Property Let AuftragsartSonder(ByVal vNewValue As Byte)
iAuftragsartSonder = vNewValue
End Property
    
Public Property Get AuftragsartVersand() As Byte
AuftragsartVersand = iAuftragsartVersand
End Property
Public Property Let AuftragsartVersand(ByVal vNewValue As Byte)
iAuftragsartVersand = vNewValue
End Property
    
Public Property Get Substitutionsvereinbarung() As Byte
Substitutionsvereinbarung = iSubstitutionsvereinbarung
End Property
Public Property Let Substitutionsvereinbarung(ByVal vNewValue As Byte)
iSubstitutionsvereinbarung = vNewValue
End Property

Public Property Get VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart() As Byte
VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart = iVerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart
End Property
Public Property Let VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart(ByVal vNewValue As Byte)
iVerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart = vNewValue
End Property

Public Property Get VerfuegbarkeitBulkVereinbart() As Byte
VerfuegbarkeitBulkVereinbart = iVerfuegbarkeitBulkVereinbart
End Property
Public Property Let VerfuegbarkeitBulkVereinbart(ByVal vNewValue As Byte)
iVerfuegbarkeitBulkVereinbart = vNewValue
End Property

Public Property Get RuecknahmeangebotVereinbart() As Byte
RuecknahmeangebotVereinbart = iRuecknahmeangebotVereinbart
End Property
Public Property Let RuecknahmeangebotVereinbart(ByVal vNewValue As Byte)
iRuecknahmeangebotVereinbart = vNewValue
End Property

Public Property Get GueltigAb$()
GueltigAb$ = iGueltigAb$
End Property
Public Property Let GueltigAb(ByVal vNewValue$)
iGueltigAb$ = vNewValue$
End Property

Public Property Get VereinbarteAuftragsKennung$(ByVal ind%)
VereinbarteAuftragsKennung = iVereinbarteAuftragsKennung(ind)
End Property
Public Property Let VereinbarteAuftragsKennung(ByVal ind%, ByVal vNewValue$)
iVereinbarteAuftragsKennung(ind) = vNewValue$
End Property

Public Property Get LieferVorgabe() As Byte
LieferVorgabe = iLieferVorgabe
End Property
Public Property Let LieferVorgabe(ByVal vNewValue As Byte)
iLieferVorgabe = vNewValue
End Property
    
Public Property Get AuftragsKennung$()
AuftragsKennung = iAuftragsKennung
End Property
Public Property Let AuftragsKennung(ByVal vNewValue$)
iAuftragsKennung = vNewValue$
End Property

Public Property Get GhCheckBestellen() As Byte
GhCheckBestellen = iGhCheckBestellen
End Property
Public Property Let GhCheckBestellen(ByVal vNewValue As Byte)
iGhCheckBestellen = vNewValue
End Property

Public Property Get GhCheckLieferVorgabe() As Byte
GhCheckLieferVorgabe = iGhCheckLieferVorgabe
End Property
Public Property Let GhCheckLieferVorgabe(ByVal vNewValue As Byte)
iGhCheckLieferVorgabe = vNewValue
End Property

Public Property Get GhCheckAuftragKennung$()
GhCheckAuftragKennung = iGhCheckAuftragKennung
End Property
Public Property Let GhCheckAuftragKennung(ByVal vNewValue$)
iGhCheckAuftragKennung = vNewValue$
End Property
''''''''''''''


Public Property Get AktRecno&()
AktRecno& = iAktRecno&
End Property

Sub GetDafueNummern(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetDafueNummern")
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
Dim h$, IniSection$

IniSection$ = Format(iLief%, "0")

h$ = "00"
l& = GetPrivateProfileString(IniSection$, "Modus", "00", h$, 3, LIFZUS_INI_DATEI)
iDirektBestKz = Val(Left$(h$, l&))

h$ = Space$(100)
l& = GetPrivateProfileString(IniSection$, "Fax", " ", h$, 101, LIFZUS_INI_DATEI)
iBestFax$ = Trim(Left$(h$, l&))

h$ = Space$(100)
l& = GetPrivateProfileString(IniSection$, "Modem", " ", h$, 101, LIFZUS_INI_DATEI)
iBestModem$ = Trim(Left$(h$, l&))

h$ = Space$(100)
l& = GetPrivateProfileString(IniSection$, "Email", " ", h$, 101, LIFZUS_INI_DATEI)
iBestMail$ = Trim(Left$(h$, l&))

h$ = Space$(100)
l& = GetPrivateProfileString(IniSection$, "AngebotWWW", " ", h$, 101, LIFZUS_INI_DATEI)
iAngebotWWW$ = Trim(Left$(h$, l&))

j% = 0
For i% = 0 To UBound(iLiefFuerHerst)
    iLiefFuerHerst(i%) = ""
Next i%
For i% = 0 To 2
    h$ = Trim(lifZusatz.LiefFuerHerst(i%))
    If (h$ <> "") Then
        lifZusatz.LiefFuerHerst(i%) = ""
        If (Asc(h$) <> 0) Then
            iLiefFuerHerst(j%) = h$
            j% = j% + 1
        End If
    End If
Next i%

h$ = Space$(100)
l& = GetPrivateProfileString(IniSection$, "LiefFuerHerst", " ", h$, 101, LIFZUS_INI_DATEI)
h$ = Trim(Left$(h$, l&))

i% = j%
Do
    ind% = InStr(h$, ",")
    If (ind% <= 0) Then
        Exit Do
    Else
        iLiefFuerHerst(i%) = Left$(h$, ind% - 1)
        i% = i% + 1
        If (i% > UBound(iLiefFuerHerst)) Then Exit Do
        h$ = Mid$(h$, ind% + 1)
    End If
Loop

Call clsError.DefErrPop
End Sub

Sub PutDafueNummern(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutDafueNummern")
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
Dim l&
Dim h$, h2$, IniSection$
    
IniSection$ = Format(iLief%, "0")

l& = WritePrivateProfileString(IniSection$, "Modus", Mid$(Str$(iDirektBestKz), 2), LIFZUS_INI_DATEI)

l& = WritePrivateProfileString(IniSection$, "Fax", iBestFax$, LIFZUS_INI_DATEI)
l& = WritePrivateProfileString(IniSection$, "Modem", iBestModem$, LIFZUS_INI_DATEI)
l& = WritePrivateProfileString(IniSection$, "Email", iBestMail$, LIFZUS_INI_DATEI)
l& = WritePrivateProfileString(IniSection$, "AngebotWWW", iAngebotWWW$, LIFZUS_INI_DATEI)
 
h$ = ""
For i% = 0 To UBound(iLiefFuerHerst)
    h2$ = Trim(iLiefFuerHerst(i%))
    If (h2$ <> "") Then h$ = h$ + h2$ + ","
Next i%
l& = WritePrivateProfileString(IniSection$, "LiefFuerHerst", h$, LIFZUS_INI_DATEI)

Call clsError.DefErrPop
End Sub

'Sub EinlesenLiefFuerHerst()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("EinlesenLiefFuerHerst")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, k%, m%, Erst%
'Dim h$
'
'ReDim AlleLiefFuerHerst(0)
'For i% = 1 To 200
'    Erst% = True
'    Get #dHandle%, i% + 1, lifZusatz
'    For j% = 0 To 2
'        h$ = lifZusatz.LiefFuerHerst(j%)
'        If (h$ = String(Len(h$), 0)) Then h$ = Space$(Len(h$))
'        h$ = UCase$(Trim(h$))
'        If (h$ <> "") Then
'            If (Erst%) Then
'                k% = UBound(AlleLiefFuerHerst) + 1
'                ReDim Preserve AlleLiefFuerHerst(k%)
'                AlleLiefFuerHerst(k%).lief = i%
'                m% = 0
'                Erst% = False
'            End If
'            AlleLiefFuerHerst(k%).herst(m%) = h$
'            m% = m% + 1
'        End If
'    Next j%
'Next i%
'
'Call clsError.DefErrPop
'End Sub

'Function GetLiefFuerHerst$(iHerst$, Optional StartRet$ = "")
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetLiefFuerHerst$")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%
'Dim ret$, h2$
'
'ret$ = StartRet$
'If (Trim$(iHerst$) <> "") Then
'    For i% = 1 To UBound(AlleLiefFuerHerst)
'        With AlleLiefFuerHerst(i%)
'            For j% = 0 To 2
'                If (.herst(j%) = iHerst$) Then
'                    h2$ = Format(.lief, "000") + ","
'                    If (InStr(ret$, h2$) <= 0) Then ret$ = ret$ + h2$
'                    Exit For
'                End If
'            Next j%
'        End With
'    Next i%
'End If
'
'GetLiefFuerHerst = ret$
'
'Call clsError.DefErrPop
'End Function



