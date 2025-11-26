VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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

Dim iLabels%
Dim iPosition%
Dim iBigSymbols%

Dim iIniDatei$
Dim iIniSection$

Const MAX_INFO_ZEILEN = 6
Const MIN_INFO_ZEILEN = 4

Private Const DefErrModul = "clsInfoBereich.cls"

Sub InitInfoBereich(iFlex As Object, IniDatei$, IniSection$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("InitToolbar")
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
Dim i%, j%, ind%
Dim l&
Dim h$, h2$, key$

Set ParentFlex = iFlex
iIniDatei$ = IniDatei$
iIniSection$ = IniSection$


For i% = 1 To MAX_INFO_ZEILEN
    For j% = 1 To 3
        iDatei$(i% - 1, j% - 1) = ""
        iKurz$(i% - 1, j% - 1) = ""
        iBezeichnung$(i% - 1, j% - 1) = ""
        
        h$ = Space$(100)
        key$ = "Info" + Format(i%, "0") + Format(j%, "0")
        l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
        h$ = Trim$(Left$(h$, l&))
        If (Len(h$) > 0) Then
            iAnzInfoZeilen% = i%
            ind% = InStr(h$, ",")
            h2$ = ""
            If (ind% > 0) Then
                h2$ = clsDat.FirstLettersUcase(Mid$(h$, ind% + 1))
                h$ = Left$(h$, ind% - 1)
            End If
            ind% = InStr(h$, ".")
            If (ind% > 0) Then
                iDatei$(i% - 1, j% - 1) = UCase$(Left$(h$, ind% - 1))
                iKurz$(i% - 1, j% - 1) = Mid$(h$, ind% + 1)
                iBezeichnung$(i% - 1, j% - 1) = h2$
            End If
        End If
    Next j%
Next i%

If (iAnzInfoZeilen% < MIN_INFO_ZEILEN) Then
    iAnzInfoZeilen% = MIN_INFO_ZEILEN
End If

'Call DefErrPop
End Sub

Function EinlesenDateiDS%(pzn$, dKurz$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("EinlesenDateiDS")
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
Dim FabsErrf%, ret%
Dim FabsRecno&
Dim SQLStr$

ret% = False

Select Case dKurz$
    Case "AST"
        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ret% = True
            Call Ast1.GetInfoRecord(FabsRecno& + 1)
        End If
    Case "ASS"
        FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ret% = True
            Call Ass1.GetInfoRecord(FabsRecno& + 1)
        End If
    Case "TAXE"
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            ret% = True
        End If
    Case "LIEF"
        Call Lif1.GetInfoRecord(Val(pzn$) + 1)
        ret% = True
End Select

EinlesenDateiDS% = ret%

'Call DefErrPop
End Function

Sub ZeigeInfoBereich(pzn$, Optional AuchWerte% = True)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ZeigeInfoBereich")
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

'Dim i%, j%, Send%, InfoInd%, DateiInd%, FeldInd%
'Dim sp&, x1&, X2&, y1&, Y2&, TextH&, TextW&, OrgFarbe&, NachLinks&
'Dim h$, tx$, txWert$

Dim i%, j%
Dim AstNeu%, AssNeu%, TaxeNeu%, LiefNeu%
Dim AstDa%, AssDa%, TaxeDa%, LiefDa%
Dim h$, txWert$, tx$

With ParentFlex
    .Redraw = False
    .GridLines = flexGridInset
    .SelectionMode = flexSelectionFree
    .Rows = iAnzInfoZeilen%
    .FixedRows = 0
    .Cols = 7
  
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
    
    If (AuchWerte% = False) Then
        AstNeu% = True
        AstDa% = False
        AssNeu% = True
        AssDa% = False
        TaxeNeu% = True
        TaxeDa% = False
        LiefNeu% = True
        LiefDa% = False
    End If

    For i% = 0 To 2
        For j% = 1 To iAnzInfoZeilen%
            h$ = iDatei$(j% - 1, i%)
            If (h$ <> "") Then
                If (h$ = "AST") Then
                    If (AstNeu% = False) Then
                        AstDa% = EinlesenDateiDS(pzn$, h$)
                        AstNeu% = True
                    End If
                    If (AstDa%) Then
                        txWert$ = Ast1.Info(iKurz$(j% - 1, i%))
                    Else
                        txWert$ = ""
                    End If
                    If (iBezeichnung$(j% - 1, i%) = "") Then
                        iBezeichnung$(j% - 1, i%) = Ast1.FeldBezeichnung(iKurz$(j% - 1, i%))
                    End If
                ElseIf (h$ = "ASS") Then
                    If (AssNeu% = False) Then
                        AssDa% = EinlesenDateiDS(pzn$, h$)
                        AssNeu% = True
                    End If
                    If (AssDa%) Then
                        txWert$ = Ass1.Info(iKurz$(j% - 1, i%))
                    Else
                        txWert$ = ""
                    End If
                    If (iBezeichnung$(j% - 1, i%) = "") Then
                        iBezeichnung$(j% - 1, i%) = Ass1.FeldBezeichnung(iKurz$(j% - 1, i%))
                    End If
                ElseIf (h$ = "TAXE") Then
                    If (TaxeNeu% = False) Then
                        TaxeDa% = EinlesenDateiDS(pzn$, h$)
                        TaxeNeu% = True
                    End If
                    If (TaxeDa%) Then
                        txWert$ = Taxe1.TaxeValue(iKurz$(j% - 1, i%))
                    Else
                        txWert$ = ""
                    End If
                    If (iBezeichnung$(j% - 1, i%) = "") Then
                        iBezeichnung$(j% - 1, i%) = iKurz$(j% - 1, i%)
                    End If
                ElseIf (h$ = "LIEF") Then
                    If (LiefNeu% = False) Then
                        LiefDa% = EinlesenDateiDS(pzn$, h$)
                        LiefNeu% = True
                    End If
                    If (LiefDa%) Then
                        txWert$ = Lif1.Info(iKurz$(j% - 1, i%))
                    Else
                        txWert$ = ""
                    End If
                    If (iBezeichnung$(j% - 1, i%) = "") Then
                        iBezeichnung$(j% - 1, i%) = Lif1.FeldBezeichnung(iKurz$(j% - 1, i%))
                    End If
                End If
            
                tx$ = clsDat.FirstLettersUcase(iBezeichnung$(j% - 1, i%))
                
                .TextMatrix(j% - 1, 2 * i% + 1) = tx$
                .TextMatrix(j% - 1, 2 * i% + 2) = txWert$
            Else
                .TextMatrix(j% - 1, 2 * i% + 1) = ""
                .TextMatrix(j% - 1, 2 * i% + 2) = ""
            End If
        Next j%
    Next i%
    .Redraw = True
End With

'Call DefErrPop
End Sub

Sub LoescheInfoBelegung(row%, col%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("LoescheInfoBelegung")
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
Dim i%, j%, k%, ZeileLeer%
Dim h$

ZeileLeer% = True
For i% = 0 To 2
    If (iDatei$(row%, i%) <> "") Then ZeileLeer% = False
Next i%
    
If (ZeileLeer% = True) Then
    If (iAnzInfoZeilen% > MIN_INFO_ZEILEN) Then
        For j% = row% To (iAnzInfoZeilen% - 2)
            For k% = 0 To 2
                iDatei$(j%, k%) = iDatei$(j% + 1, k%)
                iKurz$(j%, k%) = iKurz$(j% + 1, k%)
                iBezeichnung$(j%, k%) = iBezeichnung$(j% + 1, k%)
            Next k%
        Next j%
        For k% = 0 To 2
            iDatei$(MAX_INFO_ZEILEN - 1, k%) = ""
            iKurz$(MAX_INFO_ZEILEN - 1, k%) = ""
            iBezeichnung$(MAX_INFO_ZEILEN - 1, k%) = ""
        Next k%
        iAnzInfoZeilen% = iAnzInfoZeilen% - 1
    End If
Else
    iDatei$(row%, col%) = ""
    iKurz$(row%, col%) = ""
    iBezeichnung$(row%, col%) = ""
End If

'Call DefErrPop
End Sub

Sub InsertInfoBelegung(row%)
Dim j%, k%

If (iAnzInfoZeilen% < MAX_INFO_ZEILEN) Then
    For j% = (iAnzInfoZeilen% - 1) To row% Step -1
        For k% = 0 To 2
            iDatei$(j% + 1, k%) = iDatei$(j%, k%)
            iKurz$(j% + 1, k%) = iKurz$(j%, k%)
            iBezeichnung$(j% + 1, k%) = iBezeichnung$(j%, k%)
        Next k%
    Next j%
    For k% = 0 To 2
        iDatei$(row%, k%) = ""
        iKurz$(row%, k%) = ""
        iBezeichnung$(row%, k%) = ""
    Next k%
    iAnzInfoZeilen% = iAnzInfoZeilen% + 1
End If
        
End Sub

Sub EditInfoBelegung()
Dim row%, col%, ind%
Dim l&
Dim OrgInfoLayoutBelegung$, h$, key$

row% = ParentFlex.row
col% = (ParentFlex.col - 1) \ 2
InfoLayoutDatei$ = iDatei$(row%, col%)
InfoLayoutBelegung$ = iDatei$(row%, col%) + "." + iKurz$(row%, col%)
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
    iKurz$(row%, col%) = Mid$(InfoLayoutBelegung$, ind% + 1)
    iBezeichnung$(row%, col%) = ""
    h$ = iDatei$(row%, col%) + "." + iKurz$(row%, col%) + "," + iBezeichnung$(row%, col%)
    key$ = "Info" + Format(row% + 1, "0") + Format(col% + 1, "0")
    l& = WritePrivateProfileString(iIniSection$, key$, h$, iIniDatei$)
End If

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


