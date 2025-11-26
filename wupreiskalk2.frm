VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWuPreisKalk2
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Preiskalkulation"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   Icon            =   "wupreiskalk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdF6 
      Caption         =   "Drucken (F6)"
      Height          =   540
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   540
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdAlleKalk 
      Caption         =   "&Alle Kalkulationen akzeptieren"
      Height          =   540
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Aufschlagstabelle &editieren"
      Height          =   540
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   540
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flxWuPreisKalk 
      Height          =   1200
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   2117
      _Version        =   65541
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
Attribute VB_Name = "frmWuPreisKalk2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WUPREISKALK.FRM"

Private Sub cmdAlleKalk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAlleKalk_Click")
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
Dim i%, row%

With flxWuPreisKalk
    .redraw = False
    row% = .row
    For i% = 1 To .Rows - 1
        Call SetzePreisZeile(i%, 1)
    Next i%
    .row = row%
    .redraw = True
End With
cmdOk.SetFocus

Call DefErrPop

End Sub

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

If (ActiveControl.Name = flxWuPreisKalk.Name) Then
    Call EditSatz
Else
    PreisKalkErg% = True
    Call WuPreisKalkAlle
    Unload Me
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

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF2_Click")
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

OptionenModus% = 1
frmWuOptionen.Show 1
If (OptionenModus% = 2) Then
End If

Call DefErrPop
End Sub

Private Sub cmdF6_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF6_Click")
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
Call ActProgram.KontrollListe(1)
Call DefErrPop
End Sub

Private Sub flxWuPreisKalk_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuPreisKalk_KeyDown")
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
'If (KeyCode = vbKeyReturn) Then
'    Call EditSatz
'ElseIf (KeyCode = vbKeyF2) Then
'    cmdF2.Value = True
'End If
Call DefErrPop
End Sub

Sub flxWuPreisKalk_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxWuPreisKalk_KeyPress")
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
Dim col%

If (KeyAscii = vbKeySpace) Then
    Call SetzePreisZeile(flxWuPreisKalk.row)
End If

Call DefErrPop
End Sub

Sub SetzePreisZeile(row%, Optional typ% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetzePreisZeile")
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
Dim col%

With flxWuPreisKalk
    .row = row%
    col% = .col
    .col = 1
    .CellFontName = "Symbol"
    If (typ% Or (.TextMatrix(row%, 1) <> Chr$(214))) Then
        If (Trim(.TextMatrix(row%, 7)) <> "") Then
            .TextMatrix(row%, 6) = .TextMatrix(row%, 7)
        End If
        .TextMatrix(row%, 1) = Chr$(214)
'        cmdF6.Enabled = True
    Else
        .TextMatrix(row%, 1) = " "
        .TextMatrix(row%, 6) = .TextMatrix(row%, 21)
    End If
    .col = col%
End With

Call DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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
If (KeyCode = vbKeyF6) Then
    cmdF6.Value = True
End If
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
Dim i%, j%, spBreite%

Call wpara.InitFont(Me)

PreisKalkErg% = False

With flxWuPreisKalk
    .Cols = 22
    .Rows = 2
    .FixedRows = 1
'    .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>BM|>NR|>RM|>LM|^Ablauf|A"
    .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>AEP|>AVP|>AVP-Rund|>AVP-Kalk|<Kalkulation|>tAVP|>POS|^v|>Wg|A"
        
    Font.Bold = True
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth("X")
    .ColWidth(2) = TextWidth("Xxxxxx Xxxxxx Xxxxxx Xxxxxx")
    .ColWidth(3) = TextWidth("XXXXXX")
    .ColWidth(4) = TextWidth("XXX")
    .ColWidth(5) = TextWidth("99999.99")
    .ColWidth(6) = TextWidth("99999.99")
    .ColWidth(7) = TextWidth("99999.99")
    .ColWidth(8) = TextWidth("99999.99")
    
    .ColWidth(9) = TextWidth("Stamm-AEP + AMPV") + wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    .ColWidth(10) = 0    'TextWidth("99999.99")
    
    .ColWidth(11) = 0   'TextWidth("9999")
    .ColWidth(12) = 0   'TextWidth("XX")
    .ColWidth(13) = 0   'TextWidth("999")
    .ColWidth(14) = 0   'TextWidth("WII") + wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    .ColWidth(15) = 0
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    Font.Bold = False
    
    spBreite% = 0
    For i% = 1 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
'    If (spBreite% > .Width) Then
'        spBreite% = .Width
'    End If
'    .ColWidth(2) = .Width - spBreite%
    .Width = spBreite% + 90
    
    .Rows = 1
    Call ActProgram.PreisKalkBefuellen

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = 11 * .RowHeight(0) + 90

    .SelectionMode = flexSelectionFree
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 5
    .RowSel = .Rows - 1
    .ColSel = 6
    .CellFontBold = True
    
    .row = 1
    .col = 7
    .RowSel = .Rows - 1
    .ColSel = 9
    .CellBackColor = vbWhite
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 5
End With
    

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With cmdF2
    .Width = TextWidth(cmdAlleKalk.Caption) + 150
    .Height = wpara.ButtonY
    .Left = flxWuPreisKalk.Left
    .Top = flxWuPreisKalk.Top + flxWuPreisKalk.Height + 150 * wpara.BildFaktor
End With

With cmdAlleKalk
    .Width = cmdF2.Width
    .Height = wpara.ButtonY
'    .Left = cmdF2.Left + cmdF2.Width + 150
'    .Top = cmdF2.Top
    .Left = flxWuPreisKalk.Left
    .Top = cmdF2.Top + cmdF2.Height + 150 * wpara.BildFaktor
End With

'Me.Width = cmdF5.Left + cmdF5.Width + 2 * wpara.LinksX
Me.Width = flxWuPreisKalk.Left + flxWuPreisKalk.Width + 2 * wpara.LinksX

With cmdEsc
'    .Top = cmdF2.Top
    .Top = cmdAlleKalk.Top
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = flxWuPreisKalk.Left + flxWuPreisKalk.Width - .Width
End With

With cmdOk
    .Top = cmdEsc.Top
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = cmdEsc.Left - 150 - .Width
End With

With cmdF6
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = (Me.ScaleWidth - .Width) / 2
    .Top = cmdEsc.Top
'    .Enabled = False
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iKalk%, mw%, aufschl%
Dim KalkAvp#
Dim KalkText$
Dim h$, h2$
            
EditModus% = 4
            
EditCol% = flxWuPreisKalk.col
If (EditCol% >= 5) And (EditCol% <= 9) Then
            
    With flxWuPreisKalk
        aRow% = .row
        .row = 0
        .CellFontBold = True
        .row = aRow%
    End With
            
    Load frmEdit
    
    If (EditCol% >= 7) Then
        lInd% = Val(flxWuPreisKalk.TextMatrix(flxWuPreisKalk.row, 17))
        aRow% = 0
        With frmEdit
            .Left = flxWuPreisKalk.Left + flxWuPreisKalk.ColPos(8) + 45
            .Left = .Left + Left + wpara.FrmBorderHeight
            .Top = flxWuPreisKalk.Top + flxWuPreisKalk.RowHeight(0)  ' (flxWuPreisKalk.row - flxWuPreisKalk.TopRow + 1) * flxWuPreisKalk.RowHeight(0)
            .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
            .Width = flxWuPreisKalk.ColWidth(8) + flxWuPreisKalk.ColWidth(9)
            .Height = flxWuPreisKalk.Height - flxWuPreisKalk.RowHeight(0)
        End With
        With frmEdit.flxEdit
            .Height = frmEdit.ScaleHeight
            frmEdit.Height = .Height
            .Width = frmEdit.ScaleWidth
            .Left = 0
            .Top = 0
            
            .Rows = 0
            .Cols = 3
            
            .ColWidth(0) = flxWuPreisKalk.ColWidth(8)
            .ColWidth(1) = TextWidth("Stamm-AEP + AMPV  ")
            .ColWidth(2) = .Width - .ColWidth(0) - .ColWidth(1)
            
            .AddItem vbTab + "(freie Kalkulation)" + vbTab + Str$(0)
'            .AddItem String$(50, "-")
            
            If (flxWuPreisKalk.TextMatrix(flxWuPreisKalk.row, 12) = "v") And (AufschlagsTabelle(MAX_AUFSCHLAEGE - 1).PreisBasis <> 0) Then
                aRow% = 1
                Call AvpKalkulation(MAX_AUFSCHLAEGE, KalkAvp#, KalkText$)
                h$ = Format(KalkAvp#, "0.00") + vbTab + KalkText$ + vbTab + Str$(MAX_AUFSCHLAEGE)
                .AddItem h$
            Else
                For i% = 0 To (MAX_AUFSCHLAEGE - 2)
                    If ((i% + 1) = lInd%) Then aRow% = i% + 1
                    Call AvpKalkulation(i% + 1, KalkAvp#, KalkText$)
                    h$ = Format(KalkAvp#, "0.00") + vbTab + KalkText$ + vbTab + Str$(i% + 1)
                    .AddItem h$
                Next i%
            End If
            
            
            .row = aRow%
            .col = 0
            .ColSel = .Cols - 1
            .Visible = True
        End With
    Else
        With frmEdit
            .Left = flxWuPreisKalk.Left + flxWuPreisKalk.ColPos(EditCol%) + 45
            .Left = .Left + Left + wpara.FrmBorderHeight
            .Top = flxWuPreisKalk.Top + (flxWuPreisKalk.row - flxWuPreisKalk.TopRow + 1) * flxWuPreisKalk.RowHeight(0)
            .Top = .Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
            .Width = flxWuPreisKalk.ColWidth(EditCol%)
            .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
        End With
        With frmEdit.txtEdit
            .Width = frmEdit.ScaleWidth
    '            .Height = frmEdit.ScaleHeight
            .Left = 0
            .Top = 0
            h2$ = flxWuPreisKalk.TextMatrix(flxWuPreisKalk.row, EditCol%)
            .text = h2$
            .BackColor = vbWhite
            .Visible = True
        End With
    End If
   
    frmEdit.Show 1
            
    With flxWuPreisKalk
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
            
        If (EditErg%) Then
            If (EditCol% >= 7) Then
                ind% = InStr(EditTxt$, vbTab)
                h$ = Left$(EditTxt$, ind% - 1)
                .TextMatrix(.row, 8) = Trim$(h$)
                EditTxt$ = Mid$(EditTxt$, ind% + 1)
        
                ind% = InStr(EditTxt$, vbTab)
                h$ = Left$(EditTxt$, ind% - 1)
                .TextMatrix(.row, 9) = Trim$(h$)
                EditTxt$ = Mid$(EditTxt$, ind% + 1)
                .TextMatrix(.row, 17) = Trim$(EditTxt$)
                
                If (Trim(.TextMatrix(.row, 8)) = "") Then
                    ManuellTxt$ = .TextMatrix(.row, 2) + " " + .TextMatrix(.row, 3) + .TextMatrix(.row, 4)
                    FreiKalkPreise#(0) = CDbl(.TextMatrix(.row, 18))
                    FreiKalkPreise#(1) = CDbl(.TextMatrix(.row, 5))
                    FreiKalkPreise#(2) = CDbl(.TextMatrix(.row, 19))
                    FreiKalkMw$ = .TextMatrix(.row, 20)
                    frmFreieKalk.Show 1
                    .TextMatrix(.row, 8) = ManuellTxt$
                End If
                
                h$ = PruefeRundung()
                .TextMatrix(.row, 7) = Trim$(h$)
            Else
                .TextMatrix(.row, EditCol%) = Format(Val(EditTxt$), "0.00")
            End If
        End If
    End With
End If

Call DefErrPop
End Sub

Sub WuPreisKalkAlle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuPreisKalkAlle")
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
Dim i%, rInd%, besStatus%
Dim pzn$

Call ww.SatzLock(1)
With flxWuPreisKalk
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = Chr$(214)) Then
            .row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                ww.WuAEP = CDbl(.TextMatrix(i%, 5))
                ww.WuAVP = CDbl(.TextMatrix(i%, 6))
                Call ActProgram.SpeicherKalkPreise(rInd%, Val(.TextMatrix(i%, 17)))
            End If
        Else
            PreisKalkErg% = False
        End If
    Next i%
End With
Call ww.SatzUnLock(1)

Call DefErrPop
End Sub

Function SucheFlexZeile%(Optional BereitsGelockt% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheFlexZeile%")
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
Dim i%, row%, pos%, ret%, Max%
Dim LaufNr&
Dim pzn$, ch$

ret% = False

With flxWuPreisKalk
    row% = .row
    LaufNr& = Val(.TextMatrix(row%, 15))
    pos% = Val(Right$(.TextMatrix(row%, 16), 5))
    
    If (BereitsGelockt% = False) Then Call ww.SatzLock(1)
    ww.GetRecord (1)
    Max% = ww.erstmax
    
    ret% = SucheDateiZeile%(pos%, Max%, LaufNr&)
    
    If (ret%) Then
'        If (bek.aktivlief > 0) Then
'            Call MsgBox("Bestellsatz gesperrt!")
'            ret% = False
'        End If
    Else
        Call iMsgBox("WÜ-Satz nicht mehr vorhanden!")
    End If
    
    If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)
End With

SucheFlexZeile% = ret%

Call DefErrPop
End Function

Sub AvpKalkulation(PkInd%, KalkAvp#, KalkText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AvpKalkulation")
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
Dim iNnAep#, iStammAep#, iTaxeAep#
Dim iMw$

KalkAvp# = 0#
KalkText$ = ""

If (PkInd% > 0) Then
    With flxWuPreisKalk
        iNnAep# = CDbl(.TextMatrix(.row, 18))
        iStammAep# = CDbl(.TextMatrix(.row, 5))
        iTaxeAep# = CDbl(.TextMatrix(.row, 19))
        iMw$ = .TextMatrix(.row, 20)
        KalkAvp# = ActProgram.AvpKalkulation(PkInd%, iNnAep#, iStammAep#, iTaxeAep#, iMw$, KalkText$)
    End With
End If

Call DefErrPop
End Sub

Function PruefeRundung$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRundung$")
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
Dim OrgAvp#, KalkAvp#, RundAvp#

With flxWuPreisKalk
    OrgAvp# = CDbl(.TextMatrix(.row, 6))
    KalkAvp# = CDbl(.TextMatrix(.row, 8))
End With

RundAvp# = ActProgram.PruefeRundung(OrgAvp#, KalkAvp#)

If (RundAvp# > 0) Then
    PruefeRundung$ = Format(RundAvp#, "0.00")
Else
    PruefeRundung$ = ""
End If

Call DefErrPop
End Function


