VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmAbholer 
   Appearance      =   0  '2D
   Caption         =   "Abholerstatus"
   ClientHeight    =   6030
   ClientLeft      =   285
   ClientTop       =   660
   ClientWidth     =   6900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6900
   Begin nlCommandButton.nlCommand nlcmdsF2 
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   2040
      Width           =   855
      _ExtentX        =   3175
      _ExtentY        =   900
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF8 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   1440
      Width           =   855
      _ExtentX        =   3175
      _ExtentY        =   900
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   840
      Width           =   855
      _ExtentX        =   3175
      _ExtentY        =   900
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   360
      Width           =   855
      _ExtentX        =   3175
      _ExtentY        =   900
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   735
      Left            =   1800
      TabIndex        =   15
      Top             =   4560
      Width           =   975
      _ExtentX        =   4551
      _ExtentY        =   1058
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
      _ExtentX        =   4551
      _ExtentY        =   1058
      Caption         =   "nlCommand"
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   6000
      Picture         =   "Abholer.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   6240
      Picture         =   "Abholer.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   6480
      Picture         =   "Abholer.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtLieferschein 
      Alignment       =   2  'Zentriert
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Text            =   "LIEFERSCHEIN !"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "Abgeholt (F5)"
      Height          =   450
      Left            =   3360
      TabIndex        =   6
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox txtZusatz 
      BorderStyle     =   0  'Kein
      Height          =   255
      Index           =   0
      Left            =   1200
      MaxLength       =   24
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdF8 
      Caption         =   "Artikel-Text (F8)"
      Height          =   450
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Edit (F2)"
      Height          =   450
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdsF2 
      Caption         =   "Artikel-Status (sF2)"
      Height          =   450
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbholerEinzeln 
      Height          =   1320
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2328
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbholerGlobal 
      Height          =   1320
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2328
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   450
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbholerInfo 
      Height          =   720
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1270
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmAbholer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ANZ_SICHTBAR_ZEILEN% = 6

Dim OrgRows%

Private Const DefErrModul = "ABHOLER.FRM"

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdOk_Click")
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
Dim i%, ind%, KisteFertig%, MaxInfoInd%
Dim lRecs&, NewId&
Dim txt$

If (txtZusatz(0).Visible) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        If (AbholerMdb%) Then
            With flxAbholerInfo
                .TextMatrix(.row, 1) = txtZusatz(0).text
            
                If (Trim(ActiveControl.text) = "") Then
                    If (iNewLine) Then
                        nlcmdOk.SetFocus
                    Else
                        cmdOk.SetFocus
                    End If
                Else
                    If (.row < (.Rows - 1)) Then
                        .row = .row + 1
                    End If
                    If (.row > (.TopRow + ANZ_SICHTBAR_ZEILEN - 1)) Then
                        .TopRow = .TopRow + 1
                    End If
                    txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
                    txtZusatz(0).text = .TextMatrix(.row, 1)
                End If
            End With
            With txtZusatz(0)
                .SelStart = 0
                .SelLength = Len(.text)
            End With
        Else
            ind% = ActiveControl.Index
            If (Trim(ActiveControl.text) = "") Or (ind% >= 5) Then
                If (iNewLine) Then
                    nlcmdOk.SetFocus
                Else
                    cmdOk.SetFocus
                End If
            Else
                txtZusatz(ind% + 1).SetFocus
            End If
        End If
    Else
        If (AbholerSQL) Then
            MaxInfoInd% = 0
            With flxAbholerInfo
                If (OrgRows% > 0) Then
                    MaxInfoInd% = Val(.TextMatrix(OrgRows% - 1, 2))
                End If
                
                For i% = OrgRows% To (.Rows - 1)
                    txt$ = Trim$(.TextMatrix(i%, 1))
                    If (txt$ <> "") Then
                        MaxInfoInd% = MaxInfoInd% + 1
                        
                        lRecs = 0
                        SQLStr = "INSERT INTO AbholerInfo (Status) VALUES (" + CStr(5) + ")"
                        Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
                        If (lRecs = 1) Then
                            On Error Resume Next
                            AbholerInfoAdoRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            SQLStr = "SELECT SCOPE_IDENTITY() AS NewID"
                            AbholerInfoAdoRec.Open SQLStr$, AbholerConn
                            If Not (AbholerInfoAdoRec.EOF) Then
                                NewId& = clsOpTool.CheckNullLong(AbholerInfoAdoRec!NewId)
                            End If
                        
                            AbholerInfoAdoRec.Close
                            SQLStr$ = "SELECT * FROM AbholerInfo WHERE Id=" + CStr(NewId)
                            AbholerInfoAdoRec.Open SQLStr, AbholerConn, adOpenDynamic, adLockOptimistic
                        
                            AbholerInfoAdoRec!AbholerDetailId = AbholerDetailAdoRec!Id     'AbholerDetailId&
                            AbholerInfoAdoRec!InfoInd = MaxInfoInd%
                            
                            AbholerInfoAdoRec!Status = 2
                            AbholerInfoAdoRec!AngelegtAm = .TextMatrix(OrgRows%, 0)
                            AbholerInfoAdoRec!AngelegtVon = BesorgerBenutzer%
                            AbholerInfoAdoRec!AngelegtComputer = Val(Para1.user)
                            AbholerInfoAdoRec!AngelegtLiRe = 0
                            AbholerInfoAdoRec!text = txt$
                            
                            AbholerInfoAdoRec.Update
                        End If
                    End If
                Next i%
                Call clsOpTool.ErhoeheCounter
            End With
        ElseIf (AbholerMdb%) Then
            Set AbholerInfoRec = AbholerDB.OpenRecordset("AbholerInfo", dbOpenTable)
            AbholerInfoRec.Index = "Id"
            
            MaxInfoInd% = 0
            With flxAbholerInfo
                If (OrgRows% > 0) Then
                    MaxInfoInd% = Val(.TextMatrix(OrgRows% - 1, 2))
                End If
                
                For i% = OrgRows% To (.Rows - 1)
                    txt$ = Trim$(.TextMatrix(i%, 1))
                    If (txt$ <> "") Then
                        MaxInfoInd% = MaxInfoInd% + 1
                        
                        AbholerInfoRec.AddNew
                        
                        AbholerInfoRec!AbholerDetailId = AbholerDetailRec!Id ' AbholerDetailId&
                        AbholerInfoRec!InfoInd = MaxInfoInd%
                        
                        AbholerInfoRec!Status = 2
                        AbholerInfoRec!AngelegtAm = .TextMatrix(OrgRows%, 0)
                        AbholerInfoRec!AngelegtVon = BesorgerBenutzer%
                        AbholerInfoRec!AngelegtComputer = Val(Para1.user)
                        AbholerInfoRec!AngelegtLiRe = 0
                        AbholerInfoRec!text = txt$
                        
                        AbholerInfoRec.Update
                    End If
                Next i%
                Call clsOpTool.ErhoeheCounter
            End With
        Else
            For i% = 0 To 5 '5 to 10
                Kiste1.InfoText(i% + 3) = RTrim(txtZusatz(i%))
            Next i%
            Kiste1.PutInhalt (flxAbholerGlobal.row - 5)
        End If

        For i% = 0 To 5
            txtZusatz(i%).Visible = False
        Next i%
        flxAbholerGlobal.Enabled = True
        flxAbholerEinzeln.Enabled = True
        If (iNewLine) Then
            nlcmdF2.Visible = True
            nlcmdF5.Visible = True
            nlcmdF8.Visible = True
            nlcmdsF2.Visible = True
            nlcmdEsc.Visible = False
'            nlcmdOk.Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.Width - nlcmdOk.Width
            nlcmdOk.Left = nlcmdsF2.Left + nlcmdsF2.Width + 150
            nlcmdOk.Cancel = True
        Else
            cmdF2.Visible = True
            cmdF5.Visible = True
            cmdF8.Visible = True
            cmdsF2.Visible = True
            cmdEsc.Visible = False
            cmdOk.Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.Width - cmdOk.Width
            cmdOk.Cancel = True
        End If
        Call clsDialog.ZeigeAbholerEinzeln(flxAbholerGlobal.row - 5)
        flxAbholerGlobal.SetFocus
    End If
Else
    Unload Me
End If

Call clsError.DefErrPop
End Sub

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
Dim i%

If (txtZusatz(0).Visible) Then
    For i% = 0 To 5
        txtZusatz(i%).Visible = False
    Next i%
    flxAbholerGlobal.Enabled = True
    flxAbholerEinzeln.Enabled = True
    
    If (iNewLine) Then
        nlcmdF2.Visible = True
        nlcmdF5.Visible = True
        nlcmdF8.Visible = True
        nlcmdsF2.Visible = True
        nlcmdEsc.Visible = False
'        nlcmdOk.Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.Width - nlcmdOk.Width
        nlcmdOk.Left = nlcmdsF2.Left + nlcmdsF2.Width + 150
        nlcmdOk.Cancel = True
    Else
        cmdF2.Visible = True
        cmdF5.Visible = True
        cmdF8.Visible = True
        cmdsF2.Visible = True
        cmdEsc.Visible = False
        cmdOk.Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.Width - cmdOk.Width
        cmdOk.Cancel = True
    End If
    Call clsDialog.ZeigeAbholerEinzeln(flxAbholerGlobal.row - 5)
    flxAbholerGlobal.SetFocus
Else
    Unload Me
End If

Call clsError.DefErrPop
End Sub

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF2_Click")
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

If (AbholerMdb%) Then
    With flxAbholerInfo
'        .HighLight = flexHighlightNever
        
        .FillStyle = flexFillRepeat
        .row = .FixedRows
        .col = 0    '.FixedCols
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = wPara1.FarbeDunklerBereich
        .FillStyle = flexFillSingle
        
        For i% = (.Rows - 1) To 1 Step -1
            If (Trim(.TextMatrix(i%, 1)) = "") Then
                .Rows = i%
            Else
                Exit For
            End If
        Next i%
        
        OrgRows% = .Rows
        .Rows = .Rows + 20
        .row = OrgRows%
        
        .TextMatrix(.row, 0) = Format(Now, "DD.MM.YYYY HH:MM")
    
'        If (ActBenutzer% > 0) Then
'            .TextMatrix(.row, 1) = Trim(para.Personal(ActBenutzer%))
'        End If
'
'        .TextMatrix(.row, 2) = Format(ProgrammTypStr$(1), "0")
        
        .TopRow = .FixedRows
        Do
            If (.row > (.TopRow + ANZ_SICHTBAR_ZEILEN - 2)) Then
                .TopRow = .TopRow + 1
            Else
                Exit Do
            End If
        Loop
        
        .col = .FixedCols
        .ColSel = .Cols - 1
    
        txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
        txtZusatz(0).Left = .Left + .ColPos(1) + 45
        txtZusatz(0).Height = .RowHeight(1)
        txtZusatz(0).Width = .ColWidth(1) - wPara1.FrmScrollHeight - 90
    End With
    
    With txtZusatz(0)
        .Visible = True
        .ZOrder 0
        .SetFocus
    
        .text = ""
        .SelStart = 0
        .SelLength = Len(.text)
    End With
Else
    Call ZeigeTextBoxen
End If

flxAbholerGlobal.Enabled = False
flxAbholerEinzeln.Enabled = False
If (iNewLine) Then
    nlcmdF2.Visible = False
    nlcmdF5.Visible = False
    nlcmdF8.Visible = False
    nlcmdsF2.Visible = False
    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300
    
    nlcmdOk.Cancel = False
    
    nlcmdEsc.Cancel = True
    nlcmdEsc.Visible = True
Else
    cmdF2.Visible = False
    cmdF5.Visible = False
    cmdF8.Visible = False
    cmdsF2.Visible = False
    cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
    
    cmdOk.Cancel = False
    
    cmdEsc.Cancel = True
    cmdEsc.Visible = True
End If

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
Dim h$, h2$

If (AbholerSQL) Then
'    AbholerDetailRec.Edit
    If (AbholerDetailAdoRec!Status < 4) Then
        If (InStr(Para1.Benutz, "j") > 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(AbholerDetailAdoRec!pzn)
            If (TaxeAdoDBok) Then
                Dim tAdoRec2 As New ADODB.Recordset
                tAdoRec2.Open SQLStr, TaxeAdoDB1.ActiveConn
                If (TaxeAdoRec.EOF = False) Then
                    If (tAdoRec2!BTMKz) Then
                        Call clsDialog.MessageBox("Achtung: Einträge im BTM-Buch nicht vergessen!", vbInformation, "BTM-Buch")
'                        Call clsError.DefErrPop: Exit Sub
                    End If
                End If
            End If
        End If
        
        AbholerDetailAdoRec!Status = 4
        AbholerDetailAdoRec!AbgeholtAm = Now
        AbholerDetailAdoRec!AbgeholtBei = BesorgerBenutzer%
        Call clsDialog.SpeicherBezug
        h$ = "ABGEHOLT"
        h2$ = "Fertig"
    Else
        AbholerDetailAdoRec!Status = 3
        AbholerDetailAdoRec!ErledigtAm = Now
        AbholerDetailAdoRec!ErledigtVon = BesorgerBenutzer%
        AbholerDetailAdoRec!AbgeholtAm = "01.01.1999 15:00"
        h$ = "FERTIG"
        h2$ = "Abgeholt"
    End If
    AbholerDetailAdoRec.Update
    Call clsOpTool.ErhoeheCounter
ElseIf (AbholerMdb%) Then
    AbholerDetailRec.Edit
    If (AbholerDetailRec!Status < 4) Then
        If (InStr(Para1.Benutz, "j") > 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(AbholerDetailRec!pzn)
            If (TaxeAdoDBok) Then
                Dim tAdoRec As New ADODB.Recordset
                tAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
                If (TaxeAdoRec.EOF = False) Then
                    If (tAdoRec!BTMKz) Then
                        Call clsDialog.MessageBox("Achtung: Einträge im BTM-Buch nicht vergessen!", vbInformation, "BTM-Buch")
'                        Call clsError.DefErrPop: Exit Sub
                    End If
                End If
            End If
        End If
        
        AbholerDetailRec!Status = 4
        AbholerDetailRec!AbgeholtAm = Now
        AbholerDetailRec!AbgeholtBei = BesorgerBenutzer%
        Call clsDialog.SpeicherBezug
        h$ = "ABGEHOLT"
        h2$ = "Fertig"
    Else
        AbholerDetailRec!Status = 3
        AbholerDetailRec!ErledigtAm = Now
        AbholerDetailRec!ErledigtVon = BesorgerBenutzer%
        AbholerDetailRec!AbgeholtAm = "01.01.1999 15:00"
        h$ = "FERTIG"
        h2$ = "Abgeholt"
    End If
    AbholerDetailRec.Update
    Call clsOpTool.ErhoeheCounter
Else
    If (Kiste1.Status < 4) Then
        Kiste1.Status = 4
        If (Para1.Land = "A") Then
            Call Kiste1.SpeicherDosBezug(flxAbholerGlobal.row - 5)
        End If
        h$ = "ABGEHOLT"
        h2$ = "Fertig"
    Else
        Kiste1.Status = 3
        h$ = "FERTIG"
        h2$ = "Abgeholt"
    End If
    Kiste1.PutInhalt (flxAbholerGlobal.row - 5)
End If

flxAbholerGlobal.TextMatrix(flxAbholerGlobal.row, 2) = h$
If (iNewLine) Then
    With nlcmdF5
        .Caption = h2
        .key = "F5"
    End With
Else
    cmdF5.Caption = h2$ + " (F5)"
End If

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
Dim h$, pzn$

With flxAbholerGlobal
    h$ = .TextMatrix(.row, 1)
End With

If (AbholerSQL) Then
    pzn$ = clsOpTool.PznString(AbholerDetailAdoRec!pzn)
ElseIf (AbholerMdb%) Then
    pzn$ = clsOpTool.PznString(AbholerDetailRec!pzn)
Else
    pzn$ = Kiste1.pzn
End If

Call clsDialog.ZusatzFenster(ZUSATZ_ARTIKEL, pzn$, h$)

Call clsError.DefErrPop
End Sub

Private Sub cmdsF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdsF2_Click")
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
Dim h$, pzn$

With flxAbholerGlobal
    h$ = .TextMatrix(.row, 1)
End With

If (AbholerSQL) Then
    pzn$ = clsOpTool.PznString(AbholerDetailAdoRec!pzn)
ElseIf (AbholerMdb%) Then
    pzn$ = clsOpTool.PznString(AbholerDetailRec!pzn)
Else
    pzn$ = Kiste1.pzn
End If

Call clsDialog.AnzeigeFenster(TEXT_BESTELLSTATUS, pzn$, h$)

Call clsError.DefErrPop
End Sub

Private Sub flxAbholerGlobal_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxAbholerGlobal_RowColChange")
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
Static IstAktiv%

If (IstAktiv%) Then
    Call clsError.DefErrPop: Exit Sub
End If

IstAktiv% = True

With flxAbholerGlobal
    If (.Visible) Then
        If (.row < 5) Then
            .row = 5
        End If
        Call clsDialog.ZeigeAbholerEinzeln(.row - 5)
    End If
End With

IstAktiv% = 0

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
Dim i%, wi%, MaxWi%, iAdd%, iAdd2%, x%, y%

Call wPara1.InitFont(Me)

With flxAbholerGlobal
    .Rows = 50  '15
    .Cols = 4
    .FixedCols = 1
    .SelectionMode = flexSelectionByRow
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColWidth(0) = TextWidth("o.Gebühr/Aconto     ")
    .ColWidth(1) = TextWidth("Wwwwwwwwww Wwwwwwwwww Wwwwwwwwwww  Wwwwwwwwwww   ")
    .ColWidth(2) = TextWidth("ABGEHOLT    ")
    .ColWidth(3) = 0
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
'    .Height = .RowHeight(0) * 10 + 90
    .TextMatrix(0, 0) = "angelegt von"
    .TextMatrix(1, 0) = "bei"
    .TextMatrix(2, 0) = "am"
    .TextMatrix(3, 0) = "um"
    .TextMatrix(4, 0) = "für"
    .TextMatrix(5, 0) = "INHALT"
    .row = 5
    .col = 1
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .FillStyle = flexFillRepeat
    .CellBackColor = vbWhite
    .FillStyle = flexFillSingle
    .row = 5
    .ColSel = .Cols - 1
End With
With txtLieferschein
    .BackColor = wPara1.OptipharmRot ' vbRed
    .ForeColor = vbWhite
    .Width = TextWidth(.text) + 600
    .Left = flxAbholerGlobal.Left + flxAbholerGlobal.Width - .Width - 15
    .Top = flxAbholerGlobal.Top + 15
    .ZOrder 0
End With

With flxAbholerEinzeln
    .Rows = 11
    .Cols = 3
    .FixedCols = 1
    .SelectionMode = flexSelectionByRow
    .TabStop = False
    .HighLight = flexHighlightNever
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .Width = flxAbholerGlobal.Width
    .ColWidth(0) = TextWidth("o.Gebühr/Aconto     ")
    .ColWidth(1) = .Width - .ColWidth(0)
    .ColWidth(2) = 0
'    .Top = flxAbholerGlobal.Top + flxAbholerGlobal.Height + 150
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * .Rows + 90
    .TextMatrix(0, 0) = "Rezept-Nr"
    .TextMatrix(1, 0) = "o.Gebühr/Aconto"
    .TextMatrix(2, 0) = "Bestellmenge"
    .TextMatrix(3, 0) = "Preis"
    .TextMatrix(4, 0) = "Abholer-Info"
    .row = 5
    .col = 1
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .FillStyle = flexFillRepeat
    .CellBackColor = vbWhite
    .FillStyle = flexFillSingle
    .row = 4
    .ColSel = .Cols - 1
End With
With flxAbholerInfo
    .Cols = 3
    If (AbholerMdb%) Then
        .Rows = ANZ_SICHTBAR_ZEILEN
        .Height = .RowHeight(0) * .Rows + 90
        
        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightWithFocus
        .Width = flxAbholerEinzeln.Width
        For i% = 0 To 2
            .ColAlignment(i%) = flxAbholerEinzeln.ColAlignment(i%)
            .ColWidth(i%) = flxAbholerEinzeln.ColWidth(i%)
        Next i%
        
        .Left = flxAbholerEinzeln.Left
'        .TextMatrix(0, 0) = "Abholer-Info"
        .row = 0
        .ColSel = .Cols - 1
        
        .Visible = True
    
        With flxAbholerEinzeln
            .Rows = 5
            .Height = .RowHeight(0) * .Rows + 90
'            .TextMatrix(4, 0) = "Artikel-Info"
        End With
    Else
        .Visible = False
    End If
End With

For i% = 1 To 5
    Load txtZusatz(i%)
    txtZusatz(i%).TabIndex = i% + 2
Next i%

Call clsDialog.BesorgerBefuellen

With flxAbholerGlobal
    .Height = .RowHeight(0) * .Rows + 90
End With

With flxAbholerEinzeln
    .Top = flxAbholerGlobal.Top + flxAbholerGlobal.Height + 150
End With
With flxAbholerInfo
    .Top = flxAbholerEinzeln.Top + flxAbholerEinzeln.Height - 30
'    .TextMatrix(0, 0) = "Abholer-Info"
End With


If (AbholerMdb%) Then
    cmdOk.Top = flxAbholerInfo.Top + flxAbholerInfo.Height + 210
Else
    cmdOk.Top = flxAbholerEinzeln.Top + flxAbholerEinzeln.Height + 210
End If

Me.Width = flxAbholerGlobal.Width + 2 * wPara1.LinksX + 120

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
'cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2

cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height
cmdEsc.Top = cmdOk.Top

With cmdF2
    Font.Name = wPara1.FontName(1)
    Font.Size = wPara1.FontSize(1)
    .Width = TextWidth(.Caption) + 150
    .Height = cmdOk.Height
    .Top = cmdOk.Top
    .Left = wPara1.LinksX
End With

With cmdF5
    Font.Name = wPara1.FontName(1)
    Font.Size = wPara1.FontSize(1)
    .Width = TextWidth("Abgeholt (F5)") + 150
    .Height = cmdOk.Height
    .Top = cmdOk.Top
    .Left = cmdF2.Left + cmdF2.Width + 150
End With

With cmdF8
    Font.Name = wPara1.FontName(1)
    Font.Size = wPara1.FontSize(1)
    .Width = TextWidth(.Caption) + 150
    .Height = cmdOk.Height
    .Top = cmdOk.Top
    .Left = cmdF5.Left + cmdF5.Width + 150
End With

With cmdsF2
    Font.Name = wPara1.FontName(1)
    Font.Size = wPara1.FontSize(1)
    .Width = TextWidth(.Caption) + 150
    .Height = cmdOk.Height
    .Top = cmdOk.Top
    .Left = cmdF8.Left + cmdF8.Width + 150
End With

MaxWi% = cmdsF2.Left + cmdsF2.Width + 450 + cmdOk.Width
If (MaxWi% < flxAbholerGlobal.Width) Then
    MaxWi% = flxAbholerGlobal.Width
Else
    wi% = MaxWi% - flxAbholerGlobal.Width
    flxAbholerGlobal.Width = MaxWi%
    flxAbholerGlobal.ColWidth(2) = flxAbholerGlobal.ColWidth(2) + wi%
    flxAbholerEinzeln.Width = MaxWi%
    flxAbholerEinzeln.ColWidth(1) = flxAbholerEinzeln.ColWidth(1) + wi%
    flxAbholerInfo.Width = MaxWi%
    flxAbholerInfo.ColWidth(1) = flxAbholerInfo.ColWidth(1) + wi%
End If
Me.Width = MaxWi% + 2 * wPara1.LinksX + 120

If (BesorgerModus% = 0) Then
    cmdOk.Caption = "OK"
    cmdOk.default = True
    cmdOk.Cancel = True
    cmdOk.Visible = True
    cmdEsc.Visible = False
'    cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
    cmdOk.Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.Width - cmdOk.Width
Else
    cmdOk.Caption = "Fertig"
    cmdOk.default = True
    cmdOk.Cancel = False
    cmdOk.Visible = True
    cmdEsc.Cancel = True
    cmdEsc.Visible = True
    cmdOk.Left = (Me.ScaleWidth - (cmdOk.Width * 2 + 300)) / 2
    cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
End If


For i% = 0 To 5
    With txtZusatz(i%)
'        .Top = flxZvusatz.Top + i * flxZusatz.RowHeight(1)
        .Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.ColPos(1) + 45
        .Height = flxAbholerEinzeln.RowHeight(1)
        .Width = flxAbholerEinzeln.ColWidth(1) - 90
        .Visible = False
    End With
Next i%




Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight


If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxAbholerGlobal
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
    With txtLieferschein
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    With flxAbholerEinzeln
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
    With flxAbholerInfo
'        .ScrollBars = flexScrollBarNone
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
'        .Top = .Top + iAdd
        .Top = flxAbholerEinzeln.Top + flxAbholerEinzeln.Height + 30
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    cmdF2.Top = cmdF2.Top + 2 * iAdd
    cmdF5.Top = cmdF5.Top + 2 * iAdd
    cmdF8.Top = cmdF8.Top + 2 * iAdd
    cmdsF2.Top = cmdsF2.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxAbholerGlobal.Top = flxAbholerGlobal.Top + iAdd2
    flxAbholerEinzeln.Top = flxAbholerEinzeln.Top + iAdd2
    flxAbholerInfo.Top = flxAbholerInfo.Top + iAdd2
    txtLieferschein.Top = txtLieferschein.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    cmdF2.Top = cmdF2.Top + iAdd2
    cmdF5.Top = cmdF5.Top + iAdd2
    cmdF8.Top = cmdF8.Top + iAdd2
    cmdsF2.Top = cmdsF2.Top + iAdd2
    
    For i% = 0 To 5
        With txtZusatz(i%)
            .Left = flxAbholerEinzeln.Left + flxAbholerEinzeln.ColPos(1) + 45
        End With
    Next i%

    Height = Height + iAdd2

    With nlcmdOk
        .Init
        .Left = cmdOk.Left
        If (AbholerMdb%) Then
            .Top = flxAbholerInfo.Top + flxAbholerInfo.Height + 600 * iFaktorY
        Else
            .Top = flxAbholerEinzeln.Top + flxAbholerEinzeln.Height + 600 * iFaktorY
        End If
        .Top = .Top + iAdd
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdF2
        .Init
        .Left = cmdF2.Left
        .Top = nlcmdOk.Top
        .Caption = cmdF2.Caption
        .TabIndex = cmdF2.TabIndex
        .Enabled = cmdF2.Enabled
        .Visible = True 'cmdF2.Visible
        .AutoSize = True
    End With
    cmdF2.Visible = False

    With nlcmdF5
        .Init
        .Left = nlcmdF2.Left + nlcmdF2.Width + 150 * iFaktorX
        .Top = nlcmdOk.Top
        .AutoSize = True
        .Caption = "Abgeholt (F5)"
        .TabIndex = cmdF5.TabIndex
        .Enabled = cmdF5.Enabled
        .Visible = True 'cmdF5.Visible
        .AutoSize = 0
        .Caption = cmdF5.Caption
    End With
    cmdF5.Visible = False
    
    With nlcmdF8
        .Init
        .Left = nlcmdF5.Left + nlcmdF5.Width + 150 * iFaktorX
        .Top = nlcmdOk.Top
        .Caption = cmdF8.Caption
        .TabIndex = cmdF8.TabIndex
        .Enabled = cmdF8.Enabled
        .Visible = True 'cmdF5.Visible
        .AutoSize = True
    End With
    cmdF8.Visible = False

    With nlcmdsF2
        .Init
        .Left = nlcmdF8.Left + nlcmdF8.Width + 150 * iFaktorX
        .Top = nlcmdOk.Top
        .Caption = cmdsF2.Caption
        .TabIndex = cmdsF2.TabIndex
        .Enabled = cmdsF2.Enabled
        .Visible = True 'cmdF5.Visible
        .AutoSize = True
    End With
    cmdsF2.Visible = False

    If (BesorgerModus% = 0) Then
'        nlcmdOk.Caption = "OK"
        nlcmdOk.default = True
        nlcmdOk.Cancel = True
        nlcmdOk.Visible = True
        nlcmdEsc.Visible = False
'        nlcmdOk.Left = nlcmdsF2.Left + nlcmdsF2.Width + 150 * iFaktorX
        nlcmdOk.Left = Me.ScaleWidth - nlcmdOk.Width - 300 * iFaktorX
    Else
'        nlcmdOk.Caption = "Fertig"
        nlcmdOk.default = True
        nlcmdOk.Cancel = False
        nlcmdOk.Visible = True
        nlcmdEsc.Cancel = True
        nlcmdEsc.Visible = True
        nlcmdOk.Left = (Me.ScaleWidth - (nlcmdOk.Width * 2 + 300 * iFaktorX)) / 2
        nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300 * iFaktorX
    End If

'    Me.Width = nlcmdOk.Left + nlcmdOk.Width + 600 * iFaktorX
    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + iAdd2

    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxAbholerGlobal.Left - iAdd) / Screen.TwipsPerPixelX, (flxAbholerGlobal.Top - iAdd) / Screen.TwipsPerPixelY, (flxAbholerGlobal.Left + flxAbholerGlobal.Width + iAdd) / Screen.TwipsPerPixelX, (flxAbholerInfo.Top + flxAbholerInfo.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
    nlcmdF5.Visible = False
    nlcmdF8.Visible = False
    nlcmdsF2.Visible = False
End If

    
Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

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
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
    RoundRect hdc, (flxAbholerGlobal.Left - iAdd) / Screen.TwipsPerPixelX, (flxAbholerGlobal.Top - iAdd) / Screen.TwipsPerPixelY, (flxAbholerGlobal.Left + flxAbholerGlobal.Width + iAdd) / Screen.TwipsPerPixelX, (flxAbholerInfo.Top + flxAbholerInfo.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    txtLieferschein.BackColor = wPara1.OptipharmRot ' vbRed

    Call Form_Resize
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
Dim ind%

If (iNewLine) Then
    If (KeyCode = vbKeyF5) And (nlcmdF5.Enabled) Then
        nlcmdF5.Value = True
    ElseIf (KeyCode = vbKeyF8) Then
        nlcmdF8.Value = True
    ElseIf (KeyCode = vbKeyF2) And (Shift And vbShiftMask) Then
        nlcmdsF2.Value = True
    ElseIf (KeyCode = vbKeyF2) Then
        nlcmdF2.Value = True
    ElseIf (KeyCode = vbKeyDown) Then
        If (ActiveControl.Name = txtZusatz(0).Name) Then
            If (AbholerMdb%) Then
                With flxAbholerInfo
                    .TextMatrix(.row, 1) = txtZusatz(0).text
                    If (.row < (.Rows - 1)) Then
                        .row = .row + 1
                    End If
                    If (.row > (.TopRow + ANZ_SICHTBAR_ZEILEN - 1)) Then
                        .TopRow = .TopRow + 1
                    End If
                    txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
                    txtZusatz(0).text = .TextMatrix(.row, 1)
                End With
                With txtZusatz(0)
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            Else
                ind% = ActiveControl.Index
                If (ind% < 5) Then
                    txtZusatz(ind% + 1).SetFocus
                End If
            End If
            KeyCode = 0
        End If
    ElseIf (KeyCode = vbKeyUp) Then
        If (ActiveControl.Name = txtZusatz(0).Name) Then
            If (AbholerMdb%) Then
                With flxAbholerInfo
                    .TextMatrix(.row, 1) = txtZusatz(0).text
                    If (.row > OrgRows%) Then
                        .row = .row - 1
                    End If
                    If (.row < .TopRow) Then
                        .TopRow = .row
                    End If
                    txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
                    txtZusatz(0).text = .TextMatrix(.row, 1)
                End With
                With txtZusatz(0)
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            Else
                ind% = ActiveControl.Index
                If (ind% > 0) Then
                    txtZusatz(ind% - 1).SetFocus
                End If
            End If
            KeyCode = 0
        End If
    End If
Else
    If (KeyCode = vbKeyF5) And (cmdF5.Enabled) Then
        cmdF5.Value = True
    ElseIf (KeyCode = vbKeyF8) Then
        cmdF8.Value = True
    ElseIf (KeyCode = vbKeyF2) And (Shift And vbShiftMask) Then
        cmdsF2.Value = True
    ElseIf (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyDown) Then
        If (ActiveControl.Name = txtZusatz(0).Name) Then
            If (AbholerMdb%) Then
                With flxAbholerInfo
                    .TextMatrix(.row, 1) = txtZusatz(0).text
                    If (.row < (.Rows - 1)) Then
                        .row = .row + 1
                    End If
                    If (.row > (.TopRow + ANZ_SICHTBAR_ZEILEN - 1)) Then
                        .TopRow = .TopRow + 1
                    End If
                    txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
                    txtZusatz(0).text = .TextMatrix(.row, 1)
                End With
                With txtZusatz(0)
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            Else
                ind% = ActiveControl.Index
                If (ind% < 5) Then
                    txtZusatz(ind% + 1).SetFocus
                End If
            End If
            KeyCode = 0
        End If
    ElseIf (KeyCode = vbKeyUp) Then
        If (ActiveControl.Name = txtZusatz(0).Name) Then
            If (AbholerMdb%) Then
                With flxAbholerInfo
                    .TextMatrix(.row, 1) = txtZusatz(0).text
                    If (.row > OrgRows%) Then
                        .row = .row - 1
                    End If
                    If (.row < .TopRow) Then
                        .TopRow = .row
                    End If
                    txtZusatz(0).Top = .Top + (.row - .TopRow) * .RowHeight(1) + 45
                    txtZusatz(0).text = .TextMatrix(.row, 1)
                End With
                With txtZusatz(0)
                    .SelStart = 0
                    .SelLength = Len(.text)
                End With
            Else
                ind% = ActiveControl.Index
                If (ind% > 0) Then
                    txtZusatz(ind% - 1).SetFocus
                End If
            End If
            KeyCode = 0
        End If
    End If
End If

Call clsError.DefErrPop
End Sub

Sub ZeigeTextBoxen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeTextBoxen")
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

For i% = 0 To 5
    With txtZusatz(i%)
        .Top = flxAbholerEinzeln.Top + (i% + 5) * flxAbholerEinzeln.RowHeight(1) + 45
        .Visible = True
        .ZOrder 0
    End With
Next i%
txtZusatz(0).SetFocus

Call clsError.DefErrPop
End Sub

Private Sub nlcmdShow()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("nlcmdShow")
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
        c.Show
    End If
Next
On Error GoTo DefErr

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

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdF2_Click()
Call cmdF2_Click
End Sub

Private Sub nlcmdF5_Click()
Call cmdF5_Click
End Sub

Private Sub nlcmdF8_Click()
Call cmdF8_Click
End Sub

Private Sub nlcmdsF2_Click()
Call cmdsF2_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
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

