VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBmPartner 
   Caption         =   "Lagerstand Partner-Apos"
   ClientHeight    =   3555
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4875
   Icon            =   "BmPartner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxBmPartner 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmBmPartner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AufteilungPzn$

Private Const DefErrModul = "BMPARTNER.FRM"

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOk_Click")
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
Dim iLief%

If (ActiveControl.Name = cmdOk.Name) Then
    Call BmPartnerSpeichern
    With flxBmPartner
        iLief% = Val(.TextMatrix(.row, 2))
        If (iLief% > 0) Then
            EditTxt$ = Format(iLief%, "0")
            EditErg% = True
        End If
    End With
    Unload Me
ElseIf (ActiveControl.Name = flxBmPartner.Name) Then
    Call EditSatz
End If

Call DefErrPop
End Sub

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEsc_Click")
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

Unload Me

Call DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Load")
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
Dim i%, spBreite%

Call wpara.InitFont(Me)

AufteilungPzn$ = KorrPzn$

Caption = "Aufteilung Direktbezug bearbeiten" + " - " + KorrTxt$

EditErg% = 0
EditTxt$ = ""

Font.Bold = False   ' True

With flxBmPartner
    .Cols = 10
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .FormatString = ">Sort#|>Profil#|>Lief#|<Name|>POS|>l.Lieferung|>OptBM|^?|>BM|"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = 0    'TextWidth("999999")
    .ColWidth(1) = 0    'TextWidth("999999")
    .ColWidth(2) = TextWidth("99999999")
    .ColWidth(3) = TextWidth(String(35, "X"))
    .ColWidth(4) = TextWidth("9999999")
    .ColWidth(5) = TextWidth("99.99.999999   ")
    .ColWidth(6) = TextWidth("999999999")
    .ColWidth(7) = TextWidth("XXX")
    .ColWidth(8) = TextWidth("9999999")
    .ColWidth(9) = 0
'    .ColWidth(10) = wpara.FrmScrollHeight

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 6 + 90
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

Font.Bold = False   ' True

Me.Width = flxBmPartner.Width + 2 * wpara.LinksX

With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    .Top = flxBmPartner.Top + flxBmPartner.Height + 150
    .Visible = True
End With
With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = cmdOk.Left + cmdEsc.Width + 300
    .Top = cmdOk.Top
End With

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

Call BmPartnerBefuellen

Call DefErrPop
End Sub

Private Sub BmPartnerBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BmPartnerBefuellen")
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
Dim i%, iProfilNr%, iSortNr%, OhneLief%
Dim OrgBm%, ActBm%, iSum%, iBm%, iRest%, MaxRow%, MaxRest%, iHeute%
Dim multi!, bm!
Dim h$, h2$, sName$, sLief$

'''''''''''''''''''''''''''''''''
'Partner-Apos
If (FremdPznOk%) Then
    With frmAction.flxarbeit(0)
        ActBm% = Val(.TextMatrix(.row, 5)) + Val(.TextMatrix(.row, 6)) 'BM+NM
    End With
    
    With flxBmPartner
        .Rows = 1
        
        DirektAufteilungRec.Seek "=", AufteilungPzn$
        If (DirektAufteilungRec.NoMatch = False) Then
            Do
                If (DirektAufteilungRec.EOF) Then
                    Exit Do
                End If
                If (DirektAufteilungRec!pzn <> AufteilungPzn$) Then
                    Exit Do
                End If
                
                If (DirektAufteilungRec!Lief = Lieferant%) Then
                    iSortNr% = 999
                    sName$ = ""
                    sLief$ = ""
                    iProfilNr% = DirektAufteilungRec!ProfilNr
                    If (iProfilNr% > 0) Then
                        h$ = Format(iProfilNr%, "000")
                        If (InStr(OpDirektPartner$, h$ + ",") > 0) Then
                            OpPartnerRec.Seek "=", iProfilNr%
                            If (OpPartnerRec.NoMatch = False) Then
                                iSortNr% = OpPartnerRec!IntSortNr
                                sName$ = OpPartnerRec!Name
                                sLief$ = Format(OpPartnerRec!IntLiefNr, "0")
                            Else
                                sName$ = "(" + Format(iProfilNr%, "0") + ")"
                            End If
                            
                            If (OrgBm% = 0) Then
                                OrgBm% = DirektAufteilungRec!BvGes
                                If (OrgBm% = 0) Then
                                    multi! = 1
                                Else
                                    multi! = ActBm% * 1# / OrgBm%
                                End If
                            End If
                            iSum% = iSum% + DirektAufteilungRec!bv
                            
                            If (OrgBm% = ActBm%) Then
                                bm! = DirektAufteilungRec!bv
                            Else
                                bm! = DirektAufteilungRec!bv * multi!
    '                                bm! = Int(DirektAufteilungRec!bv * multi! + 0.501)
                            End If
                            iBm% = Int(bm! + 0.501)
                            iRest% = Int(bm! * 100# + 0.501) Mod 100
                            
                            h$ = Format(iSortNr%, "0")
                            h$ = h$ + vbTab + Format(iProfilNr, "0")
                            h$ = h$ + vbTab + sLief$
                            h$ = h$ + vbTab + sName$
                            
                            FremdPznRec.Seek "=", AufteilungPzn$, iProfilNr%
                            If (FremdPznRec.NoMatch = False) Then
                                h$ = h$ + vbTab + Format(FremdPznRec!pos, "0")
                                h$ = h$ + vbTab + Format(FremdPznRec!LetztLief, "DD.MM.YY")
                                h$ = h$ + vbTab + Format(FremdPznRec!opt, "0.0")
                                h$ = h$ + vbTab
                                If (FremdPznRec!Ladenhüter) Then
                                    h$ = h$ + "?"
                                End If
                            Else
                                h$ = h$ + vbTab + vbTab + vbTab + vbTab
                            End If
                            
                            h$ = h$ + vbTab + Format(iBm%, "0")
                            h$ = h$ + vbTab + Format(iRest%, "0")
                            .AddItem h$
                        End If
                    End If
                End If
                
                DirektAufteilungRec.MoveNext
            Loop
        End If

        iSortNr% = 0
        iProfilNr = 0
        sLief$ = ""
        sName$ = "Eigenbedarf"
        If (.Rows = 0) Or (OrgBm% = 0) Then
            bm! = ActBm%
        Else
            bm! = OrgBm% - iSum%
            bm! = bm! * multi!
    '        bm% = Int(bm% * multi! + 0.501)d
        End If
        iBm% = Int(bm! + 0.501)
        iRest% = Int(bm! * 100# + 0.501) Mod 100
        
        h$ = Format(iSortNr%, "0")
        h$ = h$ + vbTab + Format(iProfilNr, "0")
        h$ = h$ + vbTab + sLief$
        h$ = h$ + vbTab + sName$
        
        FabsErrf% = ass.IndexSearch(0, AufteilungPzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ass.GetRecord (FabsRecno& + 1)
            h$ = h$ + vbTab + Format(ass.PosLag, "0")
            h$ = h$ + vbTab
            If (ass.lld = 0) Then
            Else
                h2$ = sDate(ass.lld)
                h$ = h$ + Left$(h2$, 2) + "." + Mid$(h2$, 3, 2) + "." + Right$(h2$, 2)
            End If
            h$ = h$ + vbTab + Format(ass.opt, "0.0")
    
            h$ = h$ + vbTab
            h2$ = Format(Now, "DDMMYY")
            iHeute% = iDate(h2$)
            If ((ass.lld + para.MonNBest) < iHeute%) Then
                h$ = h$ + "?"
            End If
        Else
            h$ = h$ + vbTab + vbTab + vbTab + vbTab
        End If
        
        h$ = h$ + vbTab + Format(iBm%, "0")
        h$ = h$ + vbTab + Format(iRest%, "0")
        .AddItem h$
        
'''''''''''''''''''
        Set OpPartnerRec = OpPartnerDB.OpenRecordset("PartnerProfile", dbOpenTable)
        OpPartnerRec.Index = "Unique"
        If (OpPartnerRec.RecordCount > 0) Then
            OpPartnerRec.MoveFirst
        End If
        
        Do
            If (OpPartnerRec.EOF) Then Exit Do
            
            If (OpPartnerRec!BeiDirektbezug) Then
                iSortNr% = OpPartnerRec!IntSortNr
                iProfilNr% = OpPartnerRec!ProfilNr
                sName$ = OpPartnerRec!Name
                sLief$ = Format(OpPartnerRec!IntLiefNr, "0")
                iBm% = 0
                iRest% = 0
                
                OhneLief% = True
                For i% = 1 To (.Rows - 1)
                    If (Val(.TextMatrix(i%, 1)) = iProfilNr%) Then
                        OhneLief% = False
                        Exit For
                    End If
                Next i%
                
                If (OhneLief%) Then
                    h$ = Format(iSortNr%, "0")
                    h$ = h$ + vbTab + Format(iProfilNr, "0")
                    h$ = h$ + vbTab + sLief$
                    h$ = h$ + vbTab + sName$
                    
                    FremdPznRec.Seek "=", AufteilungPzn$, iProfilNr%
                    If (FremdPznRec.NoMatch = False) Then
                        h$ = h$ + vbTab + Format(FremdPznRec!pos, "0")
                        h$ = h$ + vbTab + Format(FremdPznRec!LetztLief, "DD.MM.YY")
                        h$ = h$ + vbTab + Format(FremdPznRec!opt, "0.0")
                        If (FremdPznRec!Ladenhüter) Then
                            h$ = h$ + vbTab + "?"
                        End If
                    Else
                        h$ = h$ + vbTab + vbTab + vbTab + vbTab
                    End If
                    
                    h$ = h$ + vbTab + Format(iBm%, "0")
                    h$ = h$ + vbTab + Format(iRest%, "0")
                    .AddItem h$
                End If
            End If
            
            OpPartnerRec.MoveNext
        Loop
        
'''''''''''''''''''
        
        .row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 5
        
        .FillStyle = flexFillRepeat
        .col = 8
        .row = 0
        .ColSel = .col
        .RowSel = .Rows - 1
        .CellBackColor = vbWhite
        .FillStyle = flexFillSingle
    
        .Height = .RowHeight(0) * .Rows + 90
        .ZOrder 0
        
        Do
            iSum% = 0
            For i% = 0 To (.Rows - 1)
                iSum% = iSum% + Val(.TextMatrix(i%, 8))
            Next i%
            
            If (iSum% = ActBm%) Then
                Exit Do
            ElseIf (iSum% < ActBm%) Then
                MaxRow% = -1
                MaxRest% = 0
                For i% = 0 To (.Rows - 1)
                    iRest% = Val(.TextMatrix(i%, 9))
                    If (iRest% < 50) And (iRest% > MaxRest%) Then
                        MaxRow% = i%
                        MaxRest% = iRest%
                    End If
                Next i%
                If (MaxRow% >= 0) Then
                    .TextMatrix(MaxRow%, 8) = Format(Val(.TextMatrix(MaxRow%, 8)) + 1, "0")
                    .TextMatrix(MaxRow%, 9) = ""
                Else
                    Exit Do
                End If
            Else
                MaxRow% = -1
                MaxRest% = 99
                For i% = 0 To (.Rows - 1)
                    iRest% = Val(.TextMatrix(i%, 9))
                    If (iRest% >= 50) And (iRest% < MaxRest%) Then
                        MaxRow% = i%
                        MaxRest% = iRest%
                    End If
                Next i%
                If (MaxRow% >= 0) Then
                    .TextMatrix(MaxRow%, 8) = Format(Val(.TextMatrix(MaxRow%, 8)) - 1, "0")
                    .TextMatrix(MaxRow%, 9) = ""
                Else
                    Exit Do
                End If
            End If
        Loop
        
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
    End With
End If
    
Call DefErrPop
End Sub


Private Sub BmPartnerSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BmPartnerSpeichern")
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
Dim i%, iProfilNr%, iSortNr%
Dim OrgBm%, ActBm%, iSum%, iBm%, iRest%, MaxRow%, MaxRest%, rInd%
Dim multi!, bm!
Dim h$, h2$, sName$, sLief$

'Partner-Apos
If (FremdPznOk%) Then
    With flxBmPartner
        ActBm% = 0
        For i% = 1 To (.Rows - 1)
            ActBm% = ActBm% + Val(.TextMatrix(i%, 8))
        Next i%
        
        With frmAction.flxarbeit(0)
            iBm% = ActBm% - Val(.TextMatrix(.row, 6))
            .TextMatrix(.row, 5) = Format(iBm%, "0") 'BM+NM
            ww.SatzLock (1)
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                ww.bm = iBm%
                ww.PutRecord (rInd% + 1)
            End If
            ww.SatzUnLock (1)
        End With
        
        For i% = 1 To (.Rows - 1)
            iProfilNr% = Val(.TextMatrix(i%, 1))
            If (iProfilNr% > 0) Then
                DirektAufteilungRec.Seek "=", AufteilungPzn$, iProfilNr%
                iBm% = Val(.TextMatrix(i%, 8))
                If (DirektAufteilungRec.NoMatch = False) Then
                    DirektAufteilungRec.Edit
                    DirektAufteilungRec!BvGes = ActBm%
                    DirektAufteilungRec!bv = iBm%
                    DirektAufteilungRec.Update
                ElseIf (iBm% > 0) Then
                    DirektAufteilungRec.AddNew
                    DirektAufteilungRec!pzn = AufteilungPzn$
                    DirektAufteilungRec!ProfilNr = iProfilNr%
                    DirektAufteilungRec!Lief = Lieferant%
                    DirektAufteilungRec!BvGes = ActBm%
                    DirektAufteilungRec!bv = iBm%
                    DirektAufteilungRec.Update
                End If
            End If
        Next i%
    End With
End If
    
Call DefErrPop
End Sub



Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iKalk%, mw%, aufschl%, aCol%, iKalkModus%
Dim KalkAvp#
Dim KalkText$, Col1$
Dim h$, h2$
            
With flxBmPartner
    .col = 8
    EditCol% = 8
    
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
        
Load frmEdit

With frmEdit
    .Left = flxBmPartner.Left + flxBmPartner.ColPos(EditCol%) + 45
    .Left = .Left + Left + wpara.FrmBorderHeight
    .Top = flxBmPartner.Top + (flxBmPartner.row - flxBmPartner.TopRow + 1) * flxBmPartner.RowHeight(0)
    .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight  ' + wpara.FrmMenuHeight
    .Width = flxBmPartner.ColWidth(EditCol%)
    .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit.txtEdit
    .Width = frmEdit.ScaleWidth
'            .Height = frmEdit.ScaleHeight
    .Left = 0
    .Top = 0
    h2$ = flxBmPartner.TextMatrix(flxBmPartner.row, EditCol%)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

EditModus% = 0

frmEdit.Show 1
        
With flxBmPartner
    aRow% = .row
    .row = 0
    .CellFontBold = False
    .row = aRow%
        
    If (EditErg%) Then
        m% = Val(EditTxt$)
        .TextMatrix(.row, EditCol%) = Format(m%, "0")
    End If
    
    .col = 0
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub


