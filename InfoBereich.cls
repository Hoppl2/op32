VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoBereich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim ParentFlex As MSFlexGrid

Dim iAnzInfoZeilen%
Dim iDatei$(5, 2)
Dim iKurz$(5, 2)
Dim iBezeichnung$(5, 2)
Dim iKzBelegung$(5, 2)

Dim iMinInfoZeilen%

Dim iIniDatei$
Dim iIniSection$

Dim KlassenCheck%

Const MAX_INFO_ZEILEN = 6
Const MIN_INFO_ZEILEN = 4


Private Const DefErrModul = "INFOBEREICH.CLS"

Sub InitInfoBereich(iFlex As Object, IniDatei$, IniSection$, Optional InitTyp% = 0, Optional MinInfoZeilen% = MIN_INFO_ZEILEN)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitInfoBereich")
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
Dim i%, j%, k%, ind%, WasDa%
Dim l&
Dim h$, h2$, key$

Set ParentFlex = iFlex
iIniDatei$ = IniDatei$
iIniSection$ = IniSection$

iMinInfoZeilen% = MinInfoZeilen%


WasDa% = False
For k% = 0 To 1
    For i% = 1 To MAX_INFO_ZEILEN
        For j% = 1 To 3
            iDatei$(i% - 1, j% - 1) = ""
            iKurz$(i% - 1, j% - 1) = ""
            iBezeichnung$(i% - 1, j% - 1) = ""
            iKzBelegung$(i% - 1, j% - 1) = ""
            
            h$ = Space$(100)
            key$ = "Info" + Format(i%, "0") + Format(j%, "0")
            l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
            h$ = Trim$(Left$(h$, l&))
            If (Len(h$) > 0) Then
                WasDa% = True
                iAnzInfoZeilen% = i%
                ind% = InStr(h$, ",")
                h2$ = ""
                If (ind% > 0) Then
    '                h2$ = clsDat.FirstLettersUcase(Mid$(h$, ind% + 1))
                    h2$ = Mid$(h$, ind% + 1)
                    h$ = Left$(h$, ind% - 1)
                End If
                ind% = InStr(h$, ".")
                If (ind% > 0) Then
                    iDatei$(i% - 1, j% - 1) = UCase$(Left$(h$, ind% - 1))
                    iKurz$(i% - 1, j% - 1) = Mid$(h$, ind% + 1)
                    
                    ind% = InStr(h2$, ",")
                    If (ind% > 0) Then
                        h$ = Left$(h2$, ind% - 1)
                        h2$ = Mid$(h2$, ind% + 1)
                    Else
                        h$ = h2$
                        h2$ = ""
                    End If
                    iBezeichnung$(i% - 1, j% - 1) = h$
                    iKzBelegung$(i% - 1, j - 1) = h2$
                End If
            End If
        Next j%
    Next i%
    
    If (WasDa%) Then Exit For
        
'    If (MatchTyp% = MATCH_ARTIKEL) Then
    If (InitTyp% = 0) Then
        key$ = "Info11": h$ = "TAXE.EK,Taxe:  EK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info21": h$ = "TAXE.VK,VK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info12": h$ = "AST.AEP,Stamm:  EK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info31": h$ = "TAXE.VKalt,VKalt"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info13": h$ = "ASS.POSLAG,Lagerstand"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info22": h$ = "AST.AVP,VK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info32": h$ = "NNEK.NNAEP(4),Letzt-NNEK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info33": h$ = "NNEK.LIEF(4),Letzt-Lief"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info42": h$ = "NNEK.NNAEP(3),Vorl.-NNEK"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info43": h$ = "NNEK.LIEF(3),Vorl.-Lief"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
    ElseIf (InitTyp% = 1) Then
        key$ = "Info11": h$ = "TAXE.PZN,PZN"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info21": h$ = "TAXE.HerstellerKB,Hersteller"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info31": h$ = "TAXE.BTMKz,BTM-Kz"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info41": h$ = "TAXE.HilfsmittelKz,HM-Kz"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info12": h$ = "AST.AEP,AEPapo"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info22": h$ = "TAXE.EK,AEPtaxe"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info32": h$ = "TAXE.Zuzahlung,Zuzahlung"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info42": h$ = "TAXE.Festbetrag,Festbetrag"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info13": h$ = "AST.AVP,AVPapo"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info23": h$ = "TAXE.VK,AVPtaxe"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info33": h$ = "TAXE.VKalt,AVPalt"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
        key$ = "Info43": h$ = "TAXE.lPreisAenderung,letzte Änderung"
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
    Else
        For i% = 1 To MAX_INFO_ZEILEN
            For j% = 1 To 3
                h$ = Space$(100)
                key$ = "Info" + Format(i%, "0") + Format(j%, "0")
                l& = GetPrivateProfileString("Infobereich Bestellung", key$, " ", h$, 101, iIniDatei$)
                h$ = Trim$(Left$(h$, l&))
                If (Len(h$) > 0) Then
                    l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
                End If
            Next j%
        Next i%
    End If

Next k%

'If (iAnzInfoZeilen% < MIN_INFO_ZEILEN) Then
'    iAnzInfoZeilen% = MIN_INFO_ZEILEN
'End If
If (iAnzInfoZeilen% < iMinInfoZeilen%) Then
    iAnzInfoZeilen% = iMinInfoZeilen%
End If
If (iNewLine) Then
    If (PixelX = 1680) Then
        If (iAnzInfoZeilen% < 5) Then
            iAnzInfoZeilen% = 5
        End If
    Else
        If (iAnzInfoZeilen% < 4) Then
            iAnzInfoZeilen% = 4
        End If
    End If
End If

KlassenCheck = 0

Call clsError.DefErrPop
End Sub

Sub ZeigeInfoBereich(pzn$, AuchWerte%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeInfoBereich")
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
Dim i%, j%
Dim h$, txWert$, tx$

If (AuchWerte%) And (KlassenCheck = 0) Then
'    CheckInfoBereich
    For i% = 0 To 2
        For j% = 1 To iAnzInfoZeilen%
            h$ = iDatei$(j% - 1, i%)
            If (h$ <> "") Then
                If (clsDat.CheckKlasse(h$) = 0) Then
                    iDatei$(j% - 1, i%) = ""
                    iKurz$(j% - 1, i%) = ""
                End If
            End If
        Next j%
    Next i%
    KlassenCheck = True
End If
            
With ParentFlex
    .Redraw = False
    .GridLines = flexGridInset
    .SelectionMode = flexSelectionFree
    .Rows = iAnzInfoZeilen%
    .FixedRows = 0
    .Cols = 7
  
    If (iNewLine) Then
        .GridLines = flexGridFlat
        .BackColor = vbWhite
        For i% = 1 To 6
            If (i% Mod 2) Then
                .ColAlignment(i%) = flexAlignRightCenter
            Else
                .ColAlignment(i%) = flexAlignLeftCenter
            End If
            .FillStyle = flexFillRepeat
            .col = i%
            .ColSel = i%
            .row = 0
            .RowSel = .Rows - 1
            If (i% Mod 2) Then
                .CellBackColor = wPara1.nlFlexBackColorFixed    ' RGB(199, 176, 123)
            Else
                .CellBackColor = wPara1.nlFlexBackColorInfo     ' RGB(232, 217, 172)
                .CellFontSize = .Font.Size + 2
            End If
            .FillStyle = flexFillSingle
        Next i%
    Else
        For i% = 1 To 6
            If (i% Mod 2) Then
                .ColAlignment(i%) = flexAlignRightCenter
            Else
                .ColAlignment(i%) = flexAlignLeftCenter
                .FillStyle = flexFillRepeat
                .col = i%
                .ColSel = i%
                .row = 0
                .RowSel = .Rows - 1
                .CellBackColor = vbWhite
                .FillStyle = flexFillSingle
            End If
        Next i%
    End If
    
    For i% = 0 To 2
        For j% = 1 To iAnzInfoZeilen%
            h$ = iDatei$(j% - 1, i%)
            If (h$ = "") And (iBezeichnung$(j% - 1, i%) <> "") Then
                .row = j% - 1
                .col = 2 * i% + 1
                .CellBackColor = wPara1.OptipharmRot ' vbRed
            End If
        Next j%
    Next i%
    
    ArtNeu% = False
    AstNeu% = False
    AssNeu% = False
    TaxeNeu% = False
    NnekNeu% = False
    LiefNeu% = False
    hTaxNeu% = False
    DocMorrisNeu = False
    If (AuchWerte% = False) Then
        ArtNeu% = True
        AstNeu% = True
        AssNeu% = True
        TaxeNeu% = True
        NnekNeu% = True
        LiefNeu% = True
        hTaxNeu% = True
        DocMorrisNeu = True
    End If
    
    ArtDa% = False
    AstDa% = False
    AssDa% = False
    TaxeDa% = False
    NnekDa% = False
    LiefDa% = False
    hTaxDa% = False
    DocMorrisDa = False
    SonderPrDa% = 0


    For i% = 0 To 2
        For j% = 1 To iAnzInfoZeilen%
            h$ = iDatei$(j% - 1, i%)
            If (h$ <> "") Then
                If (h$ = "KZ") Then
                    Call clsKz.CheckKzDS(pzn$, iKzBelegung$(j% - 1, i))
                Else
                    Call clsDat.CheckDateiDS(pzn$, h$, False)
                End If
            End If
        Next j%
    Next i%
            
    For i% = 0 To 2
        For j% = 1 To iAnzInfoZeilen%
            h$ = iDatei$(j% - 1, i%)
            tx$ = iBezeichnung$(j% - 1, i%)
            If (h$ <> "") Then
                If (h$ = "KZ") Then
                    txWert$ = clsKz.HoleKzValue(iKzBelegung$(j% - 1, i), pzn$)
                Else
                    txWert$ = clsDat.HoleDateiValue(h$, iKurz$(j% - 1, i%))
                End If
            Else
                txWert$ = ""
            End If
            
            .TextMatrix(j% - 1, 2 * i% + 1) = tx$
            .TextMatrix(j% - 1, 2 * i% + 2) = txWert$
        Next j%
    Next i%
    .Redraw = True
End With

Call clsError.DefErrPop
End Sub

Sub LoescheInfoBelegung(row%, col%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoescheInfoBelegung")
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

ZeileLeer% = True
For i% = 0 To 2
    If (iDatei$(row%, i%) <> "") Then ZeileLeer% = False
Next i%
    
If (ZeileLeer% = True) Then
    If (iAnzInfoZeilen% > iMinInfoZeilen%) Then
        For j% = row% To (iAnzInfoZeilen% - 2)
            For k% = 0 To 2
                iDatei$(j%, k%) = iDatei$(j% + 1, k%)
                iKurz$(j%, k%) = iKurz$(j% + 1, k%)
                iBezeichnung$(j%, k%) = iBezeichnung$(j% + 1, k%)
                iKzBelegung(j%, k%) = iKzBelegung(j% + 1, k%)
           Next k%
        Next j%
        For k% = 0 To 2
            iDatei$(MAX_INFO_ZEILEN - 1, k%) = ""
            iKurz$(MAX_INFO_ZEILEN - 1, k%) = ""
            iBezeichnung$(MAX_INFO_ZEILEN - 1, k%) = ""
            iKzBelegung(MAX_INFO_ZEILEN - 1, k%) = ""
        Next k%
        iAnzInfoZeilen% = iAnzInfoZeilen% - 1
    End If
Else
    iDatei$(row%, col%) = ""
    iKurz$(row%, col%) = ""
    iBezeichnung$(row%, col%) = ""
    iKzBelegung(row%, col%) = ""
End If
  
Call SpeicherInfoBelegung

Call clsError.DefErrPop
End Sub

Sub InsertInfoBelegung(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InsertInfoBelegung")
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

If (iAnzInfoZeilen% < MAX_INFO_ZEILEN) Then
    For j% = (iAnzInfoZeilen% - 1) To row% Step -1
        For k% = 0 To 2
            iDatei$(j% + 1, k%) = iDatei$(j%, k%)
            iKurz$(j% + 1, k%) = iKurz$(j%, k%)
            iBezeichnung$(j% + 1, k%) = iBezeichnung$(j%, k%)
            iKzBelegung(j% + 1, k%) = iKzBelegung(j%, k%)
        Next k%
    Next j%
    For k% = 0 To 2
        iDatei$(row%, k%) = ""
        iKurz$(row%, k%) = ""
        iBezeichnung$(row%, k%) = ""
        iKzBelegung(row%, k%) = ""
    Next k%
    iAnzInfoZeilen% = iAnzInfoZeilen% + 1
    
    Call SpeicherInfoBelegung
End If
        
Call clsError.DefErrPop
End Sub

Sub EditInfoBelegung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditInfoBelegung")
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
Dim row%, col%, ind%
Dim l&
Dim OrgInfoLayoutBelegung$, h$, key$, s$

row% = ParentFlex.row
col% = (ParentFlex.col - 1) \ 2
InfoLayoutDatei$ = iDatei$(row%, col%)
If (iDatei$(row, col%) = "KZ") Then
    InfoLayoutBelegung$ = ""
    h$ = iKzBelegung$(row, col%)
    If (Right$(h$, 1) <> ",") Then h$ = h$ + ","
    Do
        If (h$ = "") Then Exit Do
        ind% = InStr(h$, ",")
        s$ = Left$(h$, ind% - 1)
        h$ = Mid$(h$, ind% + 1)
    
        s$ = iDatei$(row, col%) + "." + UCase$(s$)
        InfoLayoutBelegung$ = InfoLayoutBelegung$ + s$ + " "
    Loop
Else
    InfoLayoutBelegung$ = iDatei$(row%, col%) + "." + iKurz$(row%, col%)
End If
OrgInfoLayoutBelegung$ = InfoLayoutBelegung$
Load frmInfoLayout
With ParentFlex.Container
    frmInfoLayout.Left = .Left + (.Width - frmInfoLayout.Width) / 2
    frmInfoLayout.Top = .Top + (.Height - frmInfoLayout.Height) / 2
End With
frmInfoLayout.Show 1

If (OrgInfoLayoutBelegung$ <> InfoLayoutBelegung$) Then
    ind% = InStr(InfoLayoutBelegung$, ".")
    iDatei$(row%, col%) = Left$(InfoLayoutBelegung$, ind% - 1)
    
'    iKurz$(row%, col%) = Mid$(InfoLayoutBelegung$, ind% + 1)
    If (iDatei$(row, col%) = "KZ") Then
        iKurz$(row, col%) = "kz"
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
        iKurz$(row, col%) = "kz"
        iKzBelegung$(row, col) = h$
    Else
        iKurz$(row, col%) = Mid$(InfoLayoutBelegung$, ind% + 1)
        iKzBelegung$(row, col) = ""
    End If
    
    iBezeichnung$(row%, col%) = InfoLayoutBezeichnung$ '""
    
    h$ = iDatei$(row%, col%) + "." + iKurz$(row%, col%) + "," + iBezeichnung$(row%, col%)
    If (iKzBelegung$(row, col) <> "") Then
        h$ = h$ + "," + iKzBelegung$(row, col)
    End If
    key$ = "Info" + Format(row% + 1, "0") + Format(col% + 1, "0")
    l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
End If


Call clsError.DefErrPop
End Sub

Sub SpeicherInfoBelegung()
Dim i%, j%
Dim l&
Dim h$, key$

For i% = 1 To MAX_INFO_ZEILEN
    For j% = 1 To 3
        If (i% <= iAnzInfoZeilen%) Then
            h$ = iDatei$(i% - 1, j% - 1) + "." + iKurz$(i% - 1, j% - 1) + "," + iBezeichnung$(i% - 1, j% - 1)
        Else
            h$ = ""
        End If
        key$ = "Info" + Format(i%, "0") + Format(j%, "0")
        l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
    Next j%
Next i%

End Sub

Public Property Get Bezeichnung$(ByVal row%, ByVal col%)

Bezeichnung$ = iBezeichnung$(row%, col%)

End Property

Public Property Let Bezeichnung(ByVal row%, ByVal col%, ByVal NeuerWert$)
Dim l&
Dim key$, h$

iBezeichnung$(row%, col%) = NeuerWert$

h$ = Trim(NeuerWert$)
If (h$ <> "") Then
    h$ = iDatei$(row%, col%) + "." + iKurz$(row%, col%) + "," + iBezeichnung$(row%, col%)
End If

key$ = "Info" + Format(row% + 1, "0") + Format(col% + 1, "0")
l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)

End Property

Public Property Get AnzInfoZeilen%()

AnzInfoZeilen% = iAnzInfoZeilen%

End Property


