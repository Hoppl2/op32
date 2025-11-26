VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmWumsatz 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Umsatztabelle"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "Wumsatz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdF8 
      Caption         =   "Rx Umsatz (F8)"
      Height          =   540
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   2640
      Picture         =   "Wumsatz.frx":0442
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   2880
      Picture         =   "Wumsatz.frx":04EB
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3120
      Picture         =   "Wumsatz.frx":059F
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picASumsatz 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2640
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "Löschen (F5)"
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   540
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flxWumsatz 
      Height          =   1200
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   2117
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   0
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF8 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmWumsatz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_LIEF = 250

Dim lArray#()
Dim lArray2#()
Dim l%(2)
Dim aDatum As Date

Dim Wumsatz%
Dim buf1 As String * 8
Dim aFile$

Dim ASumsatzTexte$(9)

Dim AnzeigeModus%

Private Const DefErrModul = "WUMSATZ.FRM"

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdEsc_Click")
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

Close Wumsatz%
ReDim lArray#(1, 1, 1)
ReDim lArray2#(1, 1)
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdF5_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF5_Click")
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

Call LöscheWumsatz

Call clsError.DefErrPop
End Sub

Private Sub cmdF8_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF8_Click")
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
Dim i%, j%, ok%, row%, col%, iLief%, iVal%, tmpDiff%, aYear%, tYear%, gh%, sp%
Dim tAnz&
Dim h$, ASumsatzName$, SQLStr$, rez$, sTabelle$
Dim AbDat, BisDat
Dim ASumsatzDB As Database
Dim ASumsatzRec As Recordset
Dim Td As DAO.TableDef
Dim Fld As DAO.Field
Dim ixFld As DAO.Field
Dim Idx As DAO.Index
Dim RxArray#(2, 1 To 20, 1)
Dim tValue#, tRabatt#
Dim buf$
Dim tdatum As Date

'If (cmdF8.Caption = "Umsatz (F8)") Then
If (AnzeigeModus = 2) Then
    AnzeigeModus = 0
    
    With flxWumsatz
        .TextMatrix(16, 0) = "Umsatz/Sendung"
        .TextMatrix(17, 0) = "Retourenquote"
        .TextMatrix(18, 0) = "Prognose Umsatz"
        .TextMatrix(19, 0) = "Mindest/Schwell"
        .TextMatrix(20, 0) = "Prognose %"
    End With
    DrawTable

    cmdF5.Enabled = True
    nlcmdF5.Enabled = True
        
    cmdF8.Caption = "Rx-Umsatz (F8)"
    nlcmdF8.Caption = cmdF8.Caption
'    Me.Caption = "Umsatztabelle"
Else
    AnzeigeModus = AnzeigeModus + 1
    
    ok% = True
    If (ok%) Then
        ok% = 0
        ASumsatzName$ = "ASumsatz.mdb"
        If (Dir(ASumsatzName$) <> "") Then
            On Error Resume Next
            Err.Clear
            Set ASumsatzDB = OpenDatabase(ASumsatzName$, False, False)
            If (Err.Number = 0) Then
                'Tabelle RXumsatz
                sTabelle$ = "RXumsatz"
                On Error Resume Next
                Err.Clear
                Set Td = ASumsatzDB.TableDefs(sTabelle)
                If Err.Number = 3265 Then
                    Err.Clear
                    
                    'Tabelle RxUmsatz
                    Set Td = ASumsatzDB.CreateTableDef(sTabelle)
                    
                    Set Fld = Td.CreateField("Lieferant", dbInteger)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Datum", dbDate)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Wert", dbDouble)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Anzahl", dbLong)
                    Td.Fields.Append Fld
                    
                    ' Indizes für ASumsatz
                    Set Idx = Td.CreateIndex()
                    Idx.Name = "Lieferant"
                    Idx.Primary = False
                    Idx.Unique = False
                    Set ixFld = Idx.CreateField("Lieferant")
                    Idx.Fields.Append ixFld
                    Set ixFld = Idx.CreateField("Datum")
                    Idx.Fields.Append ixFld
                    Td.Indexes.Append Idx
                    
                    ASumsatzDB.TableDefs.Append Td
                End If
                On Error GoTo DefErr
            
                'Tabelle NonRXumsatz
                sTabelle$ = "NonRXumsatz"
                On Error Resume Next
                Err.Clear
                Set Td = ASumsatzDB.TableDefs(sTabelle)
                If Err.Number = 3265 Then
                    Err.Clear
                    
                    'Tabelle NonRxUmsatz
                    Set Td = ASumsatzDB.CreateTableDef(sTabelle)
                    
                    Set Fld = Td.CreateField("Lieferant", dbInteger)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Datum", dbDate)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Wert", dbDouble)
                    Td.Fields.Append Fld
                    Set Fld = Td.CreateField("Rabatt", dbDouble)
                    Td.Fields.Append Fld
                    
                    ' Indizes für NonRXumsatz
                    Set Idx = Td.CreateIndex()
                    Idx.Name = "Lieferant"
                    Idx.Primary = False
                    Idx.Unique = False
                    Set ixFld = Idx.CreateField("Lieferant")
                    Idx.Fields.Append ixFld
                    Set ixFld = Idx.CreateField("Datum")
                    Idx.Fields.Append ixFld
                    Td.Indexes.Append Idx
                    
                    ASumsatzDB.TableDefs.Append Td
                End If
                On Error GoTo DefErr
                '''''
            
'                Set Td = ASumsatzDB.TableDefs("RXumsatz")
'                On Error Resume Next
'                Err.Clear
'                'iVal = CheckNullInt(LagerRec!LöschKz)
'                iVal% = Td.Fields("Anzahl").Attributes
'                If Err.Number = 3265 Then
'                  Err.Clear
'                '  Set Td = LagerDB.TableDefs("Lager")
'                  Td.Fields.Append Td.CreateField("Anzahl", dbLong)
'                End If
'                Err.Clear
'                On Error GoTo DefErr
                
                ok% = True
            End If
            On Error GoTo DefErr
            
            Set ASumsatzRec = ASumsatzDB.OpenRecordset("RXumsatz", dbOpenTable)
        End If
    End If
    
    If (ok%) Then
        For i = 0 To 2
            iLief% = l%(i)
            If (AnzeigeModus = 1) Then
                SQLStr$ = "SELECT * FROM RXumsatz"
            Else
                SQLStr$ = "SELECT * FROM NonRXumsatz"
            End If
            SQLStr$ = SQLStr$ + " WHERE Lieferant =" + Str$(iLief%)
            Set ASumsatzRec = ASumsatzDB.OpenRecordset(SQLStr$)
            If Not (ASumsatzRec.EOF) Then
                Do
                    If (ASumsatzRec.EOF) Then
                        Exit Do
                    End If
                    
'                    rez$ = Format$(ASumsatzRec!AbgabeSchlüssel, "0")
'                    If (InStr("156", rez$) <> 0) Then
                        tdatum = ASumsatzRec!datum
                        tValue# = ASumsatzRec!Wert
                        If (AnzeigeModus = 1) Then
                            tAnz = clsOpTool.CheckNullLong(ASumsatzRec!Anzahl)
                            
                            If (tAnz <> 0) Then
                                tmpDiff% = 13 - Abs(DateDiff("M", aDatum, tdatum))
                                If (tmpDiff% > 0) Then
                                    RxArray#(i%, tmpDiff% + 1, 0) = RxArray#(i%, tmpDiff% + 1, 0) + tValue#
                                    RxArray#(i%, tmpDiff% + 1, 1) = RxArray#(i%, tmpDiff% + 1, 1) + tAnz
                                End If
                                    
                                aYear% = Year(aDatum)
                                tYear% = Year(tdatum)
                                If (aYear% = tYear%) Then
                                    RxArray#(i%, 15, 0) = RxArray#(i%, 15, 0) + tValue#
                                    RxArray#(i%, 15, 1) = RxArray#(i%, 15, 1) + tAnz
                                ElseIf ((aYear% - 1) = tYear%) Then
                                    RxArray#(i%, 1, 0) = RxArray#(i%, 1, 0) + tValue#
                                    RxArray#(i%, 1, 1) = RxArray#(i%, 1, 1) + tAnz
                                End If
                            End If
                        Else
                            tRabatt = clsOpTool.CheckNullLong(ASumsatzRec!Rabatt)
                            
                            If (tValue <> 0) Then
                                tmpDiff% = 13 - Abs(DateDiff("M", aDatum, tdatum))
                                If (tmpDiff% > 0) Then
                                    RxArray#(i%, tmpDiff% + 1, 0) = RxArray#(i%, tmpDiff% + 1, 0) + tValue#
                                    RxArray#(i%, tmpDiff% + 1, 1) = RxArray#(i%, tmpDiff% + 1, 1) + tRabatt
                                End If
                                    
                                aYear% = Year(aDatum)
                                tYear% = Year(tdatum)
                                If (aYear% = tYear%) Then
                                    RxArray#(i%, 15, 0) = RxArray#(i%, 15, 0) + tValue#
                                    RxArray#(i%, 15, 1) = RxArray#(i%, 15, 1) + tRabatt
                                ElseIf ((aYear% - 1) = tYear%) Then
                                    RxArray#(i%, 1, 0) = RxArray#(i%, 1, 0) + tValue#
                                    RxArray#(i%, 1, 1) = RxArray#(i%, 1, 1) + tRabatt
                                End If
                            End If
                        End If
'                    End If
                    
                    ASumsatzRec.MoveNext
                Loop
            End If
        Next i
    
        With flxWumsatz
            For i = 0 To (.Rows - 1)
                For j = 1 To (.Cols - 1)
                    .TextMatrix(i, j) = ""
                Next j
            Next i
            For i = 16 To (.Rows - 1)
                .TextMatrix(i, 0) = ""
            Next i
            For i% = 1 To 3
                gh% = l%(i% - 1)
                If (LieferantenDBok) Then
                    h = ""
                    SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(gh)
                    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                    If (LieferantenRec.RecordCount <> 0) Then
                        h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
                    End If
                    LieferantenRec.Close
                ElseIf (gh > 0) And (gh <= Lif1.AnzRec) Then
                    Lif1.GetRecord (gh% + 1)
                    h$ = Trim$(Lif1.kurz)
                End If
                
                If (h$ = String$(Len(h$), 0)) Then h$ = ""
                If (h$ = "") Then
                    h$ = "(" + Str$(gh%) + ")"
                End If
                
                sp% = (i% * 2) - 1
                If (AnzeigeModus = 1) Then
                    .TextMatrix(0, sp%) = h$ + " Rx" ' Str$(l(i%))
                    .TextMatrix(0, sp% + 1) = "Anz Rx"
                Else
                    .TextMatrix(0, sp%) = h$ + " NonRx" ' Str$(l(i%))
                    .TextMatrix(0, sp% + 1) = "Rabatt"
                End If
                    
    '            .TextMatrix(1, sp%) = Format(RxArray#(i% - 1, 1, 0), "# ### ##0.00")
                For j% = 1 To 20
                    If (RxArray#(i% - 1, j%, 0) = 0#) Then
                        .TextMatrix(j%, sp%) = " "
                    Else
                        .TextMatrix(j%, sp%) = Format(RxArray#(i% - 1, j%, 0), "# ### ##0.00")
                    End If
                Next j%
            
                sp = sp + 1
    '            .TextMatrix(1, sp%) = Format(RxArray#(i% - 1, 1, 1), "# ### ##0")
                For j% = 1 To 20
                    If (RxArray#(i% - 1, j%, 1) = 0#) Then
                        .TextMatrix(j%, sp%) = " "
                    Else
                        If (AnzeigeModus = 1) Then
                            .TextMatrix(j%, sp%) = Format(RxArray#(i% - 1, j%, 1), "# ### ##0")
                        Else
                            .TextMatrix(j%, sp%) = Format(RxArray#(i% - 1, j%, 1), "# ### ##0.00")
                        End If
                    End If
                Next j%
            Next i%
        End With
        
        cmdF5.Enabled = False
        nlcmdF5.Enabled = False
        
        If (AnzeigeModus = 1) Then
            cmdF8.Caption = "NonRx-Umsatz (F8)"
            nlcmdF8.Caption = cmdF8.Caption
        Else
            cmdF8.Caption = "Umsatz (F8)"
            nlcmdF8.Caption = cmdF8.Caption
        End If
'        Me.Caption = "Rx-Umsatztabelle"
    
        picASumsatz.Visible = False
    End If
End If

With flxWumsatz
    .row = 1
    .col = 1
    .RowSel = .row
    .ColSel = .col
    .SetFocus
End With

Call clsError.DefErrPop
End Sub

Private Sub flxWumsatz_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxWumsatz_RowColChange")
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
Dim ok%, row%, col%, iLief%
Dim h$, ASumsatzName$, SQLStr$
Dim AbDat, BisDat
Dim ASumsatzDB As Database
Dim ASumsatzRec As Recordset

'If (cmdF8.Caption = "Umsatz (F8)") Then
If (AnzeigeModus = 1) Then
Else
    ok% = 0
    With flxWumsatz
        row% = .row
        col% = .col
        If (col% Mod 2) Then
            col% = col% + 1
        End If
        If (.Redraw) And (.row >= .FixedRows) And (.col >= .FixedCols) Then
            h$ = .TextMatrix(.row, 0)
            If (Val(Right$(h$, 4)) > 2000) Then
                ok% = True
            End If
        End If
    End With
    
    If (ok%) Then
        ok% = 0
        ASumsatzName$ = "ASumsatz.mdb"
        If (Dir(ASumsatzName$) <> "") Then
            On Error Resume Next
            Err.Clear
            Set ASumsatzDB = OpenDatabase(ASumsatzName$, False, False)
            Set ASumsatzRec = ASumsatzDB.OpenRecordset("ASumsatz", dbOpenTable)
            If (Err.Number = 0) Then
                ok% = True
            End If
            On Error GoTo DefErr
        End If
    End If
    
    If (ok%) Then
        iLief% = l%((col% - 1) \ 2)
        AbDat = "01." + h$
        BisDat = DateAdd("m", 1, AbDat)
        SQLStr$ = "SELECT AbgabeSchlüssel, SUM(Wert) as Wert2 FROM AsUmsatz"
        SQLStr$ = SQLStr$ + " WHERE Lieferant =" + Str$(iLief%)
        SQLStr$ = SQLStr$ + " AND Datum >= DateValue('" + Format(AbDat, "DD.MM.YY") + "')"
        SQLStr$ = SQLStr$ + " AND Datum < DateValue('" + Format(BisDat, "DD.MM.YY") + "')"
        SQLStr$ = SQLStr$ + " GROUP BY AbgabeSchlüssel"
        SQLStr$ = SQLStr$ + " ORDER BY AbgabeSchlüssel"
        Set ASumsatzRec = ASumsatzDB.OpenRecordset(SQLStr$)
        ok% = (ASumsatzRec.RecordCount > 0)
    End If
    
    'For j% = 1 To UBound(wRec)
    '    iLief% = wRec(j%).lief
    '    iLief% = lifzus.GetWumsatzLief(iLief%)
    '
    '    h$ = wRec(j%).bdatum
    '
    '    ASumsatzRec.AddNew
    '    ASumsatzRec!Lieferant = iLief%
    '    ASumsatzRec!datum = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + ".20" + Mid$(h$, 5, 2)
    '    ASumsatzRec!Wert = wRec(j%).Wert
    '    ASumsatzRec!AbgabeSchlüssel = wRec(j%).Rabatt
    '    ASumsatzRec.Update
    'Next j%
    
    
    With picASumsatz
        .Visible = False
        If (ok%) Then
            .Font.Name = wPara1.FontName(0)
            .Font.Size = wPara1.FontSize(0)
            .Width = 2 * flxWumsatz.ColWidth(0) + 180
            .Height = (TextWidth("Äg") + 30) * 10
            .Cls
            .CurrentY = 90
            ASumsatzRec.MoveFirst
            Do
                If (ASumsatzRec.EOF) Then
                    Exit Do
                End If
                
                .CurrentX = 90
                picASumsatz.Print ASumsatzTexte$(ASumsatzRec!AbgabeSchlüssel);
                
                h$ = Format(ASumsatzRec!wert2, "0.00")
                .CurrentX = 90 + 2 * flxWumsatz.ColWidth(0) - .TextWidth(h$)
                picASumsatz.Print h$
                
                ASumsatzRec.MoveNext
            Loop
            
            .Height = .CurrentY + 150
            .Top = flxWumsatz.Top + flxWumsatz.RowPos(row%) + flxWumsatz.RowHeight(0)
            .Left = flxWumsatz.Left + flxWumsatz.ColPos(col%) + flxWumsatz.ColWidth(col%) - .Width
            .Visible = True
            
            ASumsatzDB.Close
        End If
    End With
End If
    

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyDown")
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
Dim row%, col%, ind%
Dim txt$, pzn$, erg$, h2$, SQLStr$

If (KeyCode = vbKeyReturn) Then
    If (ActiveControl.Name = flxWumsatz.Name) Then
        row% = flxWumsatz.row
        col% = flxWumsatz.col
        ind% = (col% - 1) \ 2
        If ((col% Mod 2) = 0) Then col% = col% - 1
        h2$ = flxWumsatz.TextMatrix(0, col%) + ": "
        
        If (row% < 16) Then
            If (LieferantenDBok) Then
                txt = ""
                SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(l%(ind%))
                LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                If (LieferantenRec.RecordCount <> 0) Then
                    txt$ = UCase(Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz)))
                End If
                LieferantenRec.Close
            Else
                Call Lif1.GetRecord(l%(ind%) + 1)
                txt$ = UCase(Trim$(Lif1.kurz))
            End If
            
            erg$ = clsDialog.MatchCode(1, pzn$, txt$, False, False)
            If (erg$ <> "") Then
'                l%(ind%) = LifZus1.GetWumsatzLief(val(pzn$))
                l%(ind%) = Val(pzn$)
                
                aFile$ = "WUMSATZ.DAT"
                Wumsatz% = clsDat.FileOpen(aFile$, "RW")
                buf$ = Space$(10)
                Get #Wumsatz%, 1, buf$
                For i% = 0 To 2
                    Mid$(buf$, 2 + (i% * 3), 3) = Format(l%(i%) - 1, "000")
                Next i%
                Put #Wumsatz%, 1, buf$
                Close #Wumsatz%
                DrawTable
            End If
        ElseIf (row% = 16) Then
            AktWumsatzLief% = l%(ind%)
            AktWumsatzInfo$ = "S"
            AktWumsatzTyp$ = h2$ + "Umsatz/Sendung"
            frmWumsatzInfo.Show 1
        ElseIf (row% = 18) Then
            AktWumsatzLief% = l%(ind%)
            AktWumsatzInfo$ = "U"
            AktWumsatzTyp$ = h2$ + "Prognose Umsatz"
            frmWumsatzInfo.Show 1
        ElseIf (row% = 19) Then
            If (flxWumsatz.col Mod 2) Then
                Call EditSatz
            Else
                Call clsDialog.Stammdaten(Format(l%(ind%), "000"), "04010")
            End If
        End If
    
        KeyCode = 0
    End If
ElseIf (KeyCode = vbKeyF5) Then
    If (iNewLine) Then
        If (nlcmdF5.Enabled) Then
            nlcmdF5.Value = True
        End If
    Else
        If (cmdF5.Enabled) Then
            cmdF5.Value = True
        End If
    End If
    KeyCode = 0
ElseIf (KeyCode = vbKeyF8) Then
    If (iNewLine) Then
        nlcmdF8.Value = True
    Else
        cmdF8.Value = True
    End If
    KeyCode = 0
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Load")
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
Dim i%, j%, col%, spBreite%, iAdd%, iAdd2%, x%, y%, wi%

'aFile$ = Dir("WUMSATZ.DAT")
'If (aFile$ = "") Then
'    MsgBox "Datei WUMSATZ.DAT konnte nicht gefunden werden !"
'    Call clsError.DefErrPop: Exit Sub
'End If

Call wPara1.InitFont(Me)

aDatum = Format(Date, "DD.MM.YYYY")


ASumsatzTexte$(0) = "AM ApoPflicht"
ASumsatzTexte$(1) = "AM RezPflicht"
ASumsatzTexte$(2) = "AM Non ApoPflicht"
ASumsatzTexte$(3) = "Nichtarzneimittel"
ASumsatzTexte$(4) = "AM ApoPflicht, Non TaxPflicht"
ASumsatzTexte$(5) = "AM RezPflicht, Non TaxPflicht"
ASumsatzTexte$(6) = "AM RezPflicht + Rez.Zuschlag"
ASumsatzTexte$(7) = "AM ApoPflicht + Rez.Zuschlag"
ASumsatzTexte$(8) = "Droge, Chemikalien"
ASumsatzTexte$(9) = "AM RezPflicht mit Sonderregel"

Call EinlesenWumsatz


With flxWumsatz
    .Cols = 7
    .Rows = 21
    .FixedRows = 1
    .FixedCols = 1
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX

    .FormatString = "> |> |> |> |> |> |> "
    .SelectionMode = flexSelectionFree
'    .ColWidth(0) = TextWidth("September 199999  ")
    Font.Bold = True
    .ColWidth(0) = TextWidth("Prognose Umsatz  ")
    Font.Bold = False
    For i% = 1 To 6
'        .ColWidth(i%) = TextWidth("99 999 999.99")
        .ColWidth(i%) = TextWidth("XXXXXX NonRx ")
    Next i%
    
    spBreite% = 0
    For j% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(j%)
    Next j%
    .Width = spBreite% + 90
    
    .Height = .Rows * .RowHeight(0) + 90

    .FillStyle = flexFillRepeat
    
    col% = 3
    Do
        If (col% >= .Cols) Then Exit Do
        If (iNewLine) Then
            .row = 1
        Else
            .row = 0
        End If
        .col = col%
        .RowSel = .Rows - 1
        .ColSel = col% + 1
        If (iNewLine) Then
            .CellBackColor = RGB(240, 240, 240)
        Else
            .CellBackColor = vbButtonFace
        End If
        col% = col% + 4
    Loop
    
    .row = 16
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .CellFontBold = True
    .FillStyle = flexFillSingle
    
    .TextMatrix(1, 0) = "gesamt Vorjahr"
    For i% = 2 To 14
        .TextMatrix(i%, 0) = Format(DateAdd("M", i% - 14, aDatum), "MMMM YYYY")
    Next i%
    .TextMatrix(15, 0) = "gesamt akt. Jahr"
    .TextMatrix(16, 0) = "Umsatz/Sendung"
    .TextMatrix(17, 0) = "Retourenquote"
    .TextMatrix(18, 0) = "Prognose Umsatz"
    .TextMatrix(19, 0) = "Mindest/Schwell"
    .TextMatrix(20, 0) = "Prognose %"
   
    DrawTable
    
    .row = 1
    .col = 1
    .RowSel = .row
    .ColSel = .col
End With
    
Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

Me.Width = flxWumsatz.Left + flxWumsatz.Width + 2 * wPara1.LinksX

With cmdEsc
    .Top = flxWumsatz.Top + flxWumsatz.Height + 150 * wPara1.BildFaktor
    .Width = wPara1.ButtonX%
    .Height = wPara1.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

With cmdF5
    .Width = TextWidth(.Caption) + 150
    .Height = wPara1.ButtonY
    .Left = flxWumsatz.Left
    .Top = cmdEsc.Top
End With
With cmdF8
    .Width = TextWidth(.Caption) + 150
    .Height = wPara1.ButtonY
    .Left = cmdF5.Left + cmdF5.Width + 300
    .Top = cmdEsc.Top
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wPara1.TitelY% + 90 + wPara1.FrmCaptionHeight


If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxWumsatz
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wPara1.nlFlexBackColor    'vbWhite
        .BackColorBkg = wPara1.nlFlexBackColor    'vbWhite
        .BackColorFixed = wPara1.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wPara1.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    cmdF5.Top = cmdF5.Top + 2 * iAdd
    cmdF8.Top = cmdF8.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxWumsatz.Top = flxWumsatz.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    cmdF5.Top = cmdF5.Top + iAdd2
    
    Height = Height + iAdd2

    With nlcmdEsc
        .Init
        .Left = (flxWumsatz.Left + flxWumsatz.Width - .Width)
        .Top = flxWumsatz.Top + flxWumsatz.Height + 600 * iFaktorY
        .Top = .Top + iAdd
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdF5
        .Init
        .Left = cmdF5.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdF5.Caption
        .TabIndex = cmdF5.TabIndex
        .Enabled = cmdF5.Enabled
        .Visible = True 'cmdF2.Visible
        .AutoSize = True
    End With
    cmdF5.Visible = False

    With nlcmdF8
        .Init
        .Left = nlcmdF5.Left + nlcmdF5.Width + 300
        .Top = nlcmdEsc.Top
        .Caption = cmdF8.Caption
        .TabIndex = cmdF8.TabIndex
        .Enabled = cmdF8.Enabled
        .Visible = True 'cmdF2.Visible
        .AutoSize = True
    End With
    cmdF8.Visible = False

    Me.Width = nlcmdEsc.Left + nlcmdEsc.Width + 600 * iFaktorX
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top)
'    RoundRect hdc, (flxWumsatz.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatz.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatz.Left + flxWumsatz.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatz.Top + flxWumsatz.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
Else
    nlcmdEsc.Visible = False
    nlcmdF5.Visible = False
End If

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

Call clsError.DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Paint")
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
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (Para1.Newline) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    Call wPara1.NewLineWindow(Me, nlcmdEsc.Top, False)
    RoundRect hdc, (flxWumsatz.Left - iAdd) / Screen.TwipsPerPixelX, (flxWumsatz.Top - iAdd) / Screen.TwipsPerPixelY, (flxWumsatz.Left + flxWumsatz.Width + iAdd) / Screen.TwipsPerPixelX, (flxWumsatz.Top + flxWumsatz.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Sub DrawTable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DrawTable")
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
Dim i%, j%, k%, gh%, sp%
Dim tmp#
Dim h$, SQLStr$

With flxWumsatz
    For i% = 1 To 3
        gh% = l%(i% - 1)
        
        If (LieferantenDBok) Then
            On Error Resume Next
            LieferantenRec.Close
            Err.Clear
            On Error GoTo DefErr

            SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(gh)
            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
            If (LieferantenRec.RecordCount <> 0) Then
                h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
            End If
            LieferantenRec.Close
        ElseIf (gh% > 0) And (gh% <= Lif1.AnzRec) Then
            Lif1.GetRecord (gh% + 1)
            h$ = Trim$(Lif1.kurz)
        End If
        
        If (h$ = String$(Len(h$), 0)) Then h$ = ""
        If (h$ = "") Then
            h$ = "(" + Str$(gh%) + ")"
        End If
        
        For k% = 0 To 1
            If (k% = 0) Then
                sp% = (i% * 2) - 1
                .TextMatrix(0, sp%) = h$ ' Str$(l(i%))
            Else
                sp% = sp% + 1
                .TextMatrix(0, sp%) = "rabattf."
            End If
            
            .TextMatrix(1, sp%) = Format(lArray#(k%, gh%, 1), "# ### ##0.00")
            For j% = 2 To 20
                If (lArray#(k%, gh%, j%) = 0#) Then
                    .TextMatrix(j%, sp%) = " "
                ElseIf (j% = 17) Then
                    .TextMatrix(j%, sp%) = Format(lArray#(k%, gh%, j%), "# ### ##0.00 ") + "%"
                ElseIf (j% = 20) Then
                    .TextMatrix(j%, sp%) = Format(CLng(lArray#(k%, gh%, j%)), "# ### ##0 ") + "%"
                Else
                    .TextMatrix(j%, sp%) = Format(lArray#(k%, gh%, j%), "# ### ##0.00")
                End If
            Next j%
    '        tmp# = CalcProg(gh%)
    '        .TextMatrix(18, i%) = Format(tmp# / 100, "# ##0 %")
        Next k%
    Next i%
End With

Call clsError.DefErrPop
End Sub

Function CalcProg(lief As Integer) As Double
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CalcProg")
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
Dim aDay, nDay, aMonat As Integer
Dim tmp As Double

If lArray#(lief, 16) > 0 Then
    aDay = Day(Date)
    aMonat = Month(Date)
    Select Case aMonat
    Case 1, 3, 5, 7, 8, 10, 12
        nDay = 31
    Case 4, 6, 9, 11
        nDay = 30
    Case Else
        nDay = 28
    End Select
    tmp = (nDay * lArray#(lief, 14)) / aDay
    tmp = (tmp * 100) / lArray#(lief, 16)
    CalcProg = tmp
Else
    CalcProg = 0
End If

Call clsError.DefErrPop
End Function

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSatz")
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
Dim EditRow%, EditCol%, EditInd%
Dim lRecs&
Dim dVal#
Dim h2$

EditModus% = 0

EditRow% = flxWumsatz.row   '17
EditCol% = flxWumsatz.col
EditInd% = (flxWumsatz.col - 1) \ 2

Load frmEdit2

With frmEdit2
    .Left = flxWumsatz.Left + flxWumsatz.ColPos(EditCol%)
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    If (iNewLine = 0) Then
        .Left = .Left + 45
    End If
    .Top = flxWumsatz.Top + EditRow% * flxWumsatz.RowHeight(0)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Width = flxWumsatz.ColWidth(EditCol%)
    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit2.txtEdit
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = flxWumsatz.TextMatrix(EditRow%, EditCol%)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit2.Show 1
           
If (EditErg%) Then
    dVal# = Val(EditTxt$)
    If (EditRow% = 19) Then
        If (LieferantenDBok) Then
            SQLStr = "UPDATE LieferantenZusatz SET MindestUmsatz=" + clsOpTool.uFormat(dVal#, "0.00")
            SQLStr = SQLStr + " WHERE LiefNr=" + CStr(l%(EditInd%))
'            LieferantenComm.CommandText = SQLStr$
'            LieferantenComm.CommandTimeout = 120
'            LieferantenComm.Execute
            LieferantenDB1.ActiveConn.CommandTimeout = 120
            Call LieferantenDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
        Else
            LifZus1.GetRecord (l%(EditInd%) + 1)
            LifZus1.MindestUmsatz = dVal#
            LifZus1.PutRecord (l%(EditInd%) + 1)
        End If
        
        lArray#((EditCol% + 1) Mod 2, l%(EditInd%), 19) = dVal#
        DrawTable
    Else
'        EditErg% = vbYes
'        If (dVal# < 1000#) And (dVal# > 0#) Then
'            EditErg% = MsgBox("Wollen Sie bei diesem Lieferanten" + vbCrLf _
'                + "wirklich nur " + EditTxt$ + " DM" + vbCrLf _
'                + "umsetzen ???", vbYesNo + vbQuestion, "Sind Sie sicher ?")
'        ElseIf (dVal# > 1000000#) Then
'            EditErg% = MsgBox("Wollen Sie bei diesem Lieferanten" + vbCrLf _
'                + "wirklich " + EditTxt$ + " DM" + vbCrLf _
'                + "umsetzen ???", vbYesNo + vbQuestion, "Sind Sie sicher ?")
'        End If
'
'        If (EditErg% = vbYes) Then
'            buf1 = Right$(Space$(8) + Str$(dVal#), 8)
'            Put #LSCHWELL%, l%(EditCol% - 1), buf1
'            lArray#(l%(EditCol% - 1), 16) = dVal#
'            DrawTable
'        End If
    End If
End If

Call clsError.DefErrPop
End Sub

Sub LöscheWumsatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LöscheWumsatz")
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
Dim WUMSATZ_NEU%, fehler%, i%
Dim buf$, Lösch$
Dim LöschDatum As Date, tdatum As Date

Lösch$ = "3112" + Format((Year(Date) - 2) Mod 100, "00")
Lösch$ = clsDialog.MyInputBox("Löschen bis inkl. ", "Umsatzdatei", Lösch$)
If (Lösch$ = "") Then Call clsError.DefErrPop: Exit Sub

fehler% = 0
On Error GoTo ErrorWumsatz
LöschDatum = DateValue(Left(Lösch$, 2) + "." + Mid(Lösch$, 3, 2) + "." + Mid(Lösch$, 5, 2))
On Error GoTo DefErr
If (fehler% > 0) Then Call clsError.DefErrPop: Exit Sub

MousePointer = vbHourglass

aFile$ = "WUMSATZ.DAT"
Wumsatz% = clsDat.FileOpen(aFile$, "I")

aFile$ = "WUMSATZ.NEU"
WUMSATZ_NEU% = clsDat.FileOpen(aFile$, "O")
Line Input #Wumsatz%, buf$
Print #WUMSATZ_NEU%, buf$

Do
    Line Input #Wumsatz%, buf$
    If (EOF(Wumsatz%)) Then Exit Do
    
    fehler% = 0
    On Error GoTo ErrorWumsatz
    tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
    On Error GoTo DefErr
    If (fehler% = 0) Then
        If (tdatum > LöschDatum) Then
            Print #WUMSATZ_NEU%, buf$
        End If
    End If
Loop

Close #WUMSATZ_NEU%
Close #Wumsatz%

aFile$ = "WUMSATZ.OLD"
If (Dir(aFile$) <> "") Then Kill aFile$

Name "wumsatz.dat" As "wumsatz.old"
Name "wumsatz.neu" As "wumsatz.dat"

Call EinlesenWumsatz
DrawTable

MousePointer = vbDefault

Call clsError.DefErrPop: Exit Sub
    
ErrorWumsatz:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call clsError.DefErrPop
End Sub

Sub EinlesenWumsatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EinlesenWumsatz")
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
Dim i%, j%, tLief%, tmpDiff%, ret%, aYear%, tYear%, fehler%, RabattFlag%, dAnzFields%, IstWawiMdb%
Dim tValue#
Dim buf$
Dim tdatum As Date

ReDim lArray#(1, 1 To MAX_LIEF, 1 To 20)
ReDim lArray2#(1, 1 To MAX_LIEF)


Wumsatz% = clsDat.FileOpen("WUMSATZ.DAT", "I")
Line Input #Wumsatz%, buf$
l%(0) = Val(Mid$(buf$, 2, 3)) + 1
l%(1) = Val(Mid$(buf$, 5, 3)) + 1
l%(2) = Val(Mid$(buf$, 8, 3)) + 1

Do While (EOF(Wumsatz%) = False)
    Line Input #Wumsatz%, buf$

    fehler% = 0
    On Error GoTo ErrorWumsatz
    tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
    On Error GoTo DefErr
    If (fehler% = 0) Then
        tLief% = Val(Mid$(buf$, 8, 3))
        If (tLief% <= 0) Or (tLief% > MAX_LIEF) Then fehler% = 1
    End If
    If (fehler% = 0) Then
        tValue# = Val(Mid$(buf$, 12, 9)) / 100#
        RabattFlag% = (Mid$(buf$, 21, 1) <> "*")
    
        tmpDiff% = 13 - Abs(DateDiff("M", aDatum, tdatum))
        If (tmpDiff% > 0) Then
            lArray#(0, tLief%, tmpDiff% + 1) = lArray#(0, tLief%, tmpDiff% + 1) + tValue#
            If (RabattFlag%) Then lArray#(1, tLief%, tmpDiff% + 1) = lArray#(1, tLief%, tmpDiff% + 1) + tValue#
            If (tmpDiff = 13) And (tValue < 0) Then
                lArray#(0, tLief%, 17) = lArray#(0, tLief%, 17) + Abs(tValue#)
'                If (RabattFlag%) Then lArray#(1, tLief%, 17) = lArray#(1, tLief%, 17) + Abs(tValue#)
            End If
        End If
            
        aYear% = Year(aDatum)
        tYear% = Year(tdatum)
        If (aYear% = tYear%) Then
            lArray#(0, tLief%, 15) = lArray#(0, tLief%, 15) + tValue#
            If (RabattFlag%) Then lArray#(1, tLief%, 15) = lArray#(1, tLief%, 15) + tValue#
        ElseIf ((aYear% - 1) = tYear%) Then
            lArray#(0, tLief%, 1) = lArray#(0, tLief%, 1) + tValue#
            If (RabattFlag%) Then lArray#(1, tLief%, 1) = lArray#(1, tLief%, 1) + tValue#
        End If
    End If
Loop
Close Wumsatz%

For i% = 1 To MAX_LIEF
    lArray2#(0, i%) = lArray#(0, i%, 17)
Next i%

If (WawiDBok) Then
    SQLStr$ = "SELECT * FROM WawiDat WHERE Status=2 ORDER BY Pzn"
    FabsErrf = WawiDB1.OpenRecordset(WawiAdoRec, SQLStr)
    If Not (WawiAdoRec.EOF) Then
        i% = 0
        Do
            If (WawiAdoRec.EOF) Then
                Exit Do
            End If
            
'            If (WawiRec!Status <> 2) Then
'                Exit Do
'            End If
            
            If (WawiAdoRec!IstAltLast) Then
                If ((WawiAdoRec!WuNeuLm < 0) And (WawiAdoRec!WuNeuLm <> -999)) Or ((WawiAdoRec!WuNeuRm < 0) And (WawiAdoRec!WuNeuLm = 0)) Then
                    buf$ = clsOpTool.sDate(WawiAdoRec!WuBelegDatum)
                    If (buf$ <> "") Then
                        tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
                        tmpDiff% = Abs(DateDiff("M", aDatum, tdatum))
                        If (tmpDiff = 0) Then
                            tLief = WawiAdoRec!lief
                            tValue# = WawiAdoRec!WuAEP * Abs(WawiAdoRec!WuRm)
                            lArray#(0, tLief%, 17) = lArray#(0, tLief%, 17) + Abs(tValue#)
                        End If
                    End If
                End If
            End If
            
            WawiAdoRec.MoveNext
        Loop
    End If
Else
    On Error Resume Next
    dAnzFields% = WawiDB.TableDefs("WAWIDAT").Fields.Count '+ 1
    IstWawiMdb% = (Err = 0)
    On Error GoTo DefErr
    
    If (IstWawiMdb%) Then
        Set WawiRec = WawiDB.OpenRecordset("WawiDat", dbOpenTable)
        WawiRec.Index = "Zugeordnet"
        WawiRec.LockEdits = True
    '    WawiRec.Seek ">=", 2, " "
        WawiRec.Seek ">=", 2, 0
        If (WawiRec.NoMatch = False) Then
            i% = 0
            Do
                If (WawiRec.EOF) Then
                    Exit Do
                End If
                
                If (WawiRec!Status <> 2) Then
                    Exit Do
                End If
                
                If (WawiRec!IstAltLast) Then
                    If ((WawiRec!WuNeuLm < 0) And (WawiRec!WuNeuLm <> -999)) Or ((WawiRec!WuNeuRm < 0) And (WawiRec!WuNeuLm = 0)) Then
                        buf$ = clsOpTool.sDate(WawiRec!WuBelegDatum)
                        If (buf$ <> "") Then
                            tdatum = DateValue(Left(buf$, 2) + "." + Mid(buf$, 3, 2) + "." + Mid(buf$, 5, 2))
                            tmpDiff% = Abs(DateDiff("M", aDatum, tdatum))
                            If (tmpDiff = 0) Then
                                tLief = WawiRec!lief
                                tValue# = WawiRec!WuAEP * Abs(WawiRec!WuRm)
                                lArray#(0, tLief%, 17) = lArray#(0, tLief%, 17) + Abs(tValue#)
                            End If
                        End If
                    End If
                End If
                
                WawiRec.MoveNext
            Loop
        End If
    End If
End If
    
For i% = 1 To MAX_LIEF
    With LifZus1
        .GetRecord (i% + 1)
        For j% = 0 To 1
            lArray#(j%, i%, 16) = .UmsatzProSendung(j%)
            lArray#(j%, i%, 18) = .PrognoseUmsatz(j%)
        Next j%
        
        If (lArray#(0, i%, 17) > 0) Then
            tValue# = lArray#(0, i%, 14) + lArray2#(0, i%)
            If (tValue > 0.01) Then
                lArray#(0, i%, 17) = (lArray#(0, i%, 17) / tValue#) * 100#
            Else
                lArray#(0, i%, 17) = 100
            End If
        End If
       
        lArray#(0, i%, 19) = .MindestUmsatz
        
        If (.TabTyp = 0) Then
            lArray#(1, i%, 19) = .PrognoseSchwellwert
        Else
            lArray#(1, i%, 19) = 0#
        End If
        
        For j% = 0 To 1
            If (lArray#(j%, i%, 19) > 0#) Then
                lArray#(j%, i%, 20) = (.PrognoseUmsatz(j%) / lArray#(j%, i%, 19)) * 100#
            Else
                lArray#(j%, i%, 20) = 0#
            End If
        Next j%
    End With
Next i%
    
'Call clsError.DefErrPop: Exit Sub
    
Call clsError.DefErrPop: Exit Sub
    
ErrorWumsatz:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseDown")
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
    
If (y <= wPara1.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseMove")
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
Dim c As Object

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is nlCommand) Then
        If (c.MouseOver) Then
            c.MouseOver = 0
        End If
    End If
Next
On Error GoTo DefErr

Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
If (iNewLine) And (Me.Visible) Then
    CurrentX = wPara1.NlFlexBackY
    CurrentY = (wPara1.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdF5_Click()
Call cmdF5_Click
End Sub

Private Sub nlcmdF8_Click()
Call cmdF8_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdEsc_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    End If
End If

End Sub

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub


