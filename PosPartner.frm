VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmPosPartner 
   Caption         =   "Lagerstand Partner-Apos"
   ClientHeight    =   3960
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4725
   Icon            =   "PosPartner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4725
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3840
      Picture         =   "PosPartner.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3600
      Picture         =   "PosPartner.frx":0203
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3360
      Picture         =   "PosPartner.frx":02B7
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxPosPartner 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   360
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
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmPosPartner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "POSPARTNER.FRM"

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
Dim iLief%

With flxPosPartner
    iLief% = Val(.TextMatrix(.row, 2))
    If (iLief% > 0) Then
        EditTxt$ = Format(iLief%, "0")
        EditErg% = True
    End If
End With
Unload Me

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

Unload Me

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
Dim i%, spBreite%
Dim iAdd%, iAdd2%, x%, y%, wi%

Call wPara1.InitFont(Me)

Caption = Caption + " - " + OpPartnerTxt$

EditErg% = 0
EditTxt$ = ""

Font.Bold = False   ' True

With flxPosPartner
    .Cols = 10
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .FormatString = ">Sort#|>Profil#|>Lief#|<Name|>POS|>l.Lieferung|>OptBM|^?|^|"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = 0    'TextWidth("999999")
    .ColWidth(1) = 0    'TextWidth("999999")
    .ColWidth(2) = TextWidth("999999")
    .ColWidth(3) = TextWidth(String(30, "X"))
    .ColWidth(4) = TextWidth("99999")
    .ColWidth(5) = TextWidth("99.99.9999   ")
    .ColWidth(6) = TextWidth("9999999")
    .ColWidth(7) = TextWidth("XXX")
    .ColWidth(8) = TextWidth(String(10, "X"))
    .ColWidth(9) = wPara1.FrmScrollHeight

    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 6 + 90
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

Font.Bold = False   ' True

Me.Width = flxPosPartner.Width + 2 * wPara1.LinksX

With cmdOk
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    .Top = flxPosPartner.Top + flxPosPartner.Height + 150
    .Visible = OpPartnerBestell%
End With
With cmdEsc
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
    .Left = cmdOk.Left + cmdEsc.Width + 300
    .Top = cmdOk.Top
End With

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    With flxPosPartner
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
        .Top = .Top + iAdd
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdOk.Top
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    flxPosPartner.Top = flxPosPartner.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdOk.Top
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
'        .Width = 3000
'        .Height = 600
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = flxPosPartner.Top + flxPosPartner.Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .default = cmdOk.default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
'        .Width = 3000
'        .Height = 600
        .Left = nlcmdOk.Left + .Width + 300
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450

    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
    RoundRect hdc, (flxPosPartner.Left - iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top - iAdd) / Screen.TwipsPerPixelY, (flxPosPartner.Left + flxPosPartner.Width + iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top + flxPosPartner.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

'    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
'    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

Call PosPartnerBefuellen

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
    RoundRect hdc, (flxPosPartner.Left - iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top - iAdd) / Screen.TwipsPerPixelY, (flxPosPartner.Left + flxPosPartner.Width + iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top + flxPosPartner.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub PosPartnerBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PosPartnerBefuellen")
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
Dim i%, iProfilNr%, iSortNr%
Dim h$, sName$, sLief$

With flxPosPartner
    .Rows = 1
    If (FremdPznOk%) Then
        FremdPznRec.Seek "=", Val(OpPartnerPzn$)
        If (FremdPznRec.NoMatch = False) Then
            Do
                If (FremdPznRec.EOF) Then
                    Exit Do
                End If
                If (FremdPznRec!pzn <> Val(OpPartnerPzn$)) Then
                    Exit Do
                End If
                
                iSortNr% = 999
                sName$ = ""
                sLief$ = ""
                iProfilNr% = FremdPznRec!ProfilNr
                If (iProfilNr% > 0) Then
                    OpPartnerRec.Seek "=", iProfilNr%
                    If (OpPartnerRec.NoMatch = False) Then
                        iSortNr% = OpPartnerRec!IntSortNr
                        sName$ = OpPartnerRec!Name
                        sLief$ = Format(OpPartnerRec!IntLiefNr, "0")
                    Else
                        sName$ = "(" + Format(iProfilNr%, "0") + ")"
                    End If
                End If
                
                h$ = Format(iSortNr%, "0")
                h$ = h$ + vbTab + Format(iProfilNr, "0")
                h$ = h$ + vbTab + sLief$
                h$ = h$ + vbTab + sName$
                h$ = h$ + vbTab + Format(FremdPznRec!pos, "0")
                h$ = h$ + vbTab + Format(FremdPznRec!LetztLief, "DD.MM.YY")
                h$ = h$ + vbTab + Format(FremdPznRec!opt, "0.0")
                h$ = h$ + vbTab
                If (FremdPznRec!Ladenhüter) Then
                    h$ = h$ + "?"
                End If
                h$ = h$ + vbTab
                If (OpPartnerRec!FremdVerbund) Then
                    h$ = h$ + "Ext. Verbund"
                End If
                
                .AddItem h$
                
                FremdPznRec.MoveNext
            Loop
        End If
    End If
    
    If (.Rows = 1) Then
        .AddItem vbTab + vbTab + vbTab + "keine Einträge gefunden !"
    Else
        .row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 5
    End If
    
    .row = 1
    If (.Rows > 1) And (OpPartnerInitLief% > 0) Then
        For i% = 1 To (.Rows - 1)
            If (Val(.TextMatrix(i%, 2)) = OpPartnerInitLief%) Then
                .row = i%
                Exit For
            End If
        Next i%
    End If
    .col = 0
    .ColSel = .Cols - 1
End With

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

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    End If
End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub



