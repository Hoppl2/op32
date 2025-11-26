VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmAuswahlHA 
   AutoRedraw      =   -1  'True
   Caption         =   "Auswahl der Markt-PZN"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6600
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4080
      Picture         =   "AuswahlHA.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3840
      Picture         =   "AuswahlHA.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3600
      Picture         =   "AuswahlHA.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "Löschen (F5)"
      Height          =   450
      Left            =   4560
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtTaxMuster 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Tag             =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   5
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxTaxMuster 
      Height          =   2700
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxTaxMuster 
      Height          =   2640
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
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
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblTaxMuster 
      Caption         =   "&Name: "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmAuswahlHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "AUSWAHLHA.FRM"

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

FormErg = 0
FormErgTxt = ""

Unload Me

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
Dim i%, OrgMenge%, row%, erg%, AnzInhalt%
Dim NeuMenge#
Dim txt$

If (ActiveControl.Name = txtTaxMuster.Name) Then
    Call flxTaxMusterBefuellen
Else
    row% = flxTaxMuster(0).row
    FormErg% = 1
    FormErgTxt = flxTaxMuster(0).TextMatrix(row%, 0)
    
    Unload Me
End If

Call DefErrPop
End Sub

Private Sub flxTaxMuster_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxMuster_GotFocus")
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

With flxTaxMuster(index)
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
    
'    If (index = 0) And (txtTaxMuster.Visible) Then txtTaxMuster.text = .TextMatrix(.row, 1)
End With

Call DefErrPop
End Sub

Private Sub flxTaxMuster_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxMuster_LostFocus")
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

With flxTaxMuster(index)
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

'Private Sub flxTaxMuster_RowColChange(index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("flxTaxMuster_RowColChange")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, AnzInhalt%, iFlag%, ErrNo%
'Dim SatzPtr&, TmId&
'Dim h$, h2$
'Dim sXML$, Inhaltsstoffe$, InhSt$, Gefaesse$, Gefaess$, Zuschlaege$, Zuschlag$
'
'With flxTaxMuster(1)
'    If (index = 0) And (.Visible) Then
'        .Visible = False
'        .Rows = 1
'
'        sXML = ReadFileToText(LennartzPfad + "\" + flxTaxMuster(0).TextMatrix(flxTaxMuster(0).row, 0))
'        Inhaltsstoffe = XmlAbschnitt(sXML, "Inhaltsstoffe")
'        Do
'            InhSt = XmlAbschnitt(Inhaltsstoffe, "InhSt")
'            If (InhSt = "") Then
'                Exit Do
'            End If
'
'            With TaxierRec
'                .pzn = XmlAbschnitt(InhSt, "PZN")
'                .kurz = XmlAbschnitt(InhSt, "Bezeichnung")
'                .ActMenge = xVal(XmlAbschnitt(InhSt, "Menge"))
'                .Meh = XmlAbschnitt(InhSt, "Einheit")
'                .kp = XmlAbschnitt(InhSt, "Faktor")
'                .menge = XmlAbschnitt(InhSt, "Preiskennzeichen")
'                .ActPreis = XmlAbschnitt(InhSt, "Preis")
'                .GStufe = 1
'                .flag = MAG_ANTEILIG
'            End With
'
'            .AddItem " "
'            Call ZeigeTaxierZeile(.Rows - 1)
'        Loop
'
'        Gefaesse = XmlAbschnitt(sXML, "Gefaesse")
'        Do
'            Gefaess = XmlAbschnitt(Gefaesse, "Gefaess")
'            If (Gefaess = "") Then
'                Exit Do
'            End If
'
'            With TaxierRec
'                .pzn = XmlAbschnitt(Gefaess, "PZN")
'                .kurz = XmlAbschnitt(Gefaess, "Bezeichnung")
'                .ActMenge = xVal(XmlAbschnitt(Gefaess, "Menge"))
'                .Meh = XmlAbschnitt(Gefaess, "Einheit")
'                .kp = XmlAbschnitt(Gefaess, "Faktor")
'                .menge = XmlAbschnitt(Gefaess, "Preiskennzeichen")
'                .ActPreis = XmlAbschnitt(Gefaess, "Preis")
'                .GStufe = 1
'                .flag = MAG_GEFAESS
'            End With
'
'            .AddItem " "
'            Call ZeigeTaxierZeile(.Rows - 1)
'        Loop
'
'        Zuschlaege = XmlAbschnitt(sXML, "Zuschlaege")
'        Do
'            Zuschlag = XmlAbschnitt(Zuschlaege, "Zuschlag")
'            If (Zuschlag = "") Then
'                Exit Do
'            End If
'
'            With TaxierRec
'                .pzn = XmlAbschnitt(Zuschlag, "PZN")
'                .kurz = ""  'XmlAbschnitt(Zuschlag, "Bezeichnung")
'                .ActMenge = 1   'xVal(XmlAbschnitt(Zuschlag, "Menge"))
'                .Meh = ""   'XmlAbschnitt(Zuschlag, "Einheit")
'                .kp = XmlAbschnitt(Zuschlag, "Faktor")
'                .menge = XmlAbschnitt(Zuschlag, "Preiskennzeichen")
'                .ActPreis = XmlAbschnitt(Zuschlag, "Preis")
'                .GStufe = 1
'                .flag = MAG_PREISEINGABE
'
'                If (Val(.pzn) = 4443869) Then
'                    .kurz = Left$("Qualitätszuschlag" + Space$(Len(.kurz)), Len(.kurz))
'                ElseIf (Val(.pzn) = 6460518) Then
'                    If (Val(.menge) = 70) Then
'                        .kurz = Left$("Fix-Aufschlag " + Space$(Len(.kurz)), Len(.kurz))
'                    Else
'                        .kurz = Left$("Rezeptur-Zuschlag " + Space$(Len(.kurz)), Len(.kurz))
'                        .flag = MAG_ARBEIT
'                    End If
'                ElseIf (Val(.pzn) = 2567001) Then
'                    .kurz = Left$("BTM-Gebühr" + Space$(Len(.kurz)), Len(.kurz))
'                End If
'
'            End With
'
'            .AddItem " "
'            Call ZeigeTaxierZeile(.Rows - 1)
'        Loop
'
'
'        If (.Rows < 2) Then .AddItem " "
'        .row = 1
'        .Visible = True
'
'        If (txtTaxMuster.Visible) And (flxTaxMuster(0).Visible) Then
'            If (ActiveControl.Name <> txtTaxMuster.Name) Then
'                txtTaxMuster.text = flxTaxMuster(0).TextMatrix(flxTaxMuster(0).row, 1)
'            End If
'        End If
'    End If
'End With
'
'Call DefErrPop
'End Sub

Private Sub flxTaxMuster_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxMuster_KeyPress")
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
Dim i%, row%, gef%, col%
Dim ch$, h$

ch$ = UCase$(Chr$(KeyAscii))

If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxTaxMuster(index)
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
                    .row = i%
                    gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub flxTaxMuster_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxMuster_DblClick")
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

cmdOk.Value = True

Call DefErrPop
End Sub

'Private Sub txtTaxMuster_Change()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("txtTaxMuster_Change")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, l%
'Dim h$
'
'h$ = UCase$(Trim(txtTaxMuster.text))
'l% = Len(h$)
'
'With flxTaxMuster(0)
'    For i% = 1 To (.Rows - 1)
'        If (Left$(.TextMatrix(i%, 1), l%) = h$) Then
'            .TopRow = i%
'            .row = i%
'            Exit For
'        End If
'    Next i%
'End With
'
'Call DefErrPop
'End Sub

Private Sub txtTaxMuster_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtTaxMuster_GotFocus")
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

With txtTaxMuster
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus = 1

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control

iEditModus = 1

FormErg% = False

LennartzPzn = 0

Call wpara.InitFont(Me)


lblTaxMuster.Top = wpara.TitelY
lblTaxMuster.Left = wpara.LinksX

txtTaxMuster.Left = lblTaxMuster.Left + lblTaxMuster.Width + 150
txtTaxMuster.Top = lblTaxMuster.Top

With flxTaxMuster(0)
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Id|>Preis|>Menge|Meh|<Kurz|<Dar|<HerKB|<Pzn|"
    .Rows = 1
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth(String(15, "9"))
    .ColWidth(2) = TextWidth(String(15, "A"))
    .ColWidth(3) = TextWidth(String(5, "A"))
    .ColWidth(4) = TextWidth(String(50, "A"))
    If (SollTaxierTyp = MAG_SPEZIALITAET) Then
        .ColWidth(5) = TextWidth(String(10, "A"))
        .ColWidth(6) = TextWidth(String(10, "A"))
        .ColWidth(7) = TextWidth(String(10, "A"))
    End If
    .ColWidth(8) = wpara.FrmScrollHeight
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * 20 + 90
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    
'    If (TaxMusterModus%) Then
        .Top = lblTaxMuster.Top + lblTaxMuster.Height + 300
        txtTaxMuster.Width = .ColWidth(1)
        lblTaxMuster.Visible = True
        txtTaxMuster.Visible = True
'    End If
    
    Call flxTaxMusterBefuellen
    
    .row = 1
    .col = 1
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    .col = 0
    .ColSel = .Cols - 1
End With

'With flxTaxMuster(1)
'    .Rows = 2
'    .FixedRows = 1
'    .FormatString = ">Preis|>Menge|Meh|<Kurz|<PZN||Flag|"
'    .Rows = 1
'
'    .ColWidth(0) = TextWidth("9999999.99")
'    .ColWidth(1) = TextWidth(String(8, "9"))
'    .ColWidth(2) = TextWidth(String(4, "A"))
'    .ColWidth(3) = TextWidth(String(35, "A"))
'    .ColWidth(4) = TextWidth(String(9, "9"))
'    .ColWidth(5) = wpara.FrmScrollHeight
'
'    Breite1% = 0
'    For i% = 0 To (.Cols - 1)
'        Breite1% = Breite1% + .ColWidth(i%)
'    Next i%
'    .Width = Breite1% + 90
'    .Height = .RowHeight(0) * 7 + 90
'
'    .Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + 90
'    .Left = flxTaxMuster(0).Left
'End With

Font.Bold = False   ' True

'With cmdF5
'    .Left = flxTaxMuster(0).Left
'    .Top = flxTaxMuster(1).Top
'    .Width = TextWidth(.Caption) + 300
'    .Height = wpara.ButtonY
'End With

cmdOk.Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + 300
cmdEsc.Top = cmdOk.Top

Me.Width = flxTaxMuster(0).Left + flxTaxMuster(0).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

'With cmdF5
'    .Left = flxTaxMuster(0).Left
'    .Top = cmdOk.Top
'    .Width = TextWidth(.Caption) + 300
'    .Height = wpara.ButtonY
'End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxTaxMuster(0)
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
        
'        iArbeitAnzzeilen = 20
'        .Height = .RowHeight(0) * iArbeitAnzzeilen%
    End With
    
'    With flxTaxMuster(1)
''        .ScrollBars = flexScrollBarNone
'        .BorderStyle = 0
'        .Width = .Width - 90
'        .Height = .Height - 90
'        .GridLines = flexGridFlat
'        .GridLinesFixed = flexGridFlat
'        .GridColorFixed = .GridColor
'        .BackColor = wpara.nlFlexBackColor    'vbWhite
'        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
'        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
'        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
'        .ForeColorSel = vbBlack
'
'        .Left = .Left + iAdd
'        .Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + 3 * iAdd
'    End With
    
    Width = Width + 2 * iAdd
    Height = Height + 5 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
'        .Left = (Me.ScaleWidth - 2 * .Width - 300)
        .Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

'    With nlcmdF5
'        .Init
'        .Left = cmdF5.Left
'        .Top = nlcmdOk.Top
'        .Caption = cmdF5.Caption
'        .TabIndex = cmdF5.TabIndex
'        .Enabled = cmdF5.Enabled
'        .Default = cmdF5.Default
'        .Cancel = cmdF5.Cancel
'        .Visible = True
'    End With
'    cmdF5.Visible = False

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wpara.FrmCaptionHeight + 450

    nlcmdOk.Left = (Me.ScaleWidth - (nlcmdOk.Width + nlcmdEsc.Width + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdOk.Width + 300

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxTaxMuster(0)
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With
'    With flxTaxMuster(1)
'        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'    End With

    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (TypeOf c Is ComboBox) Then
                    Call wpara.ControlBorderless(c)
                ElseIf (c.Appearance = 1) Then
                    Call wpara.ControlBorderless(c, 2, 2)
                Else
                    Call wpara.ControlBorderless(c, 1, 1)
                End If

                If (c.Enabled) Then
                    c.BackColor = vbWhite
                Else
                    c.BackColor = Me.BackColor
                End If

                With c.Container
                    .ForeColor = RGB(180, 180, 180) ' vbWhite
                    .FillStyle = vbSolid
                    .FillColor = c.BackColor

                    RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                End With
            End If
        End If
    Next
    On Error GoTo DefErr
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF5.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

With flxTaxMuster(0)
    .col = 2
    .col = 1
End With

Call DefErrPop
End Sub

Private Sub flxTaxMusterBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxMusterBefuellen")
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
Dim i%, j%, k%, AnzTm%, ok%, SuchWert%, ind%, ErrNo%
Dim h$, h2$, h3$, SuchName$, sSQL$, pzn$
Dim bOk As Boolean

'If (TaxMusterModus% = 0) Then
    SuchWert% = -1
    SuchName$ = Trim(txtTaxMuster.text) ' TaxMusterSuch$
    
    ind% = InStr(SuchName$, ",")
    If (ind%) Then
        SuchWert% = Val(Mid$(SuchName$, ind% + 1))
        SuchName$ = Left$(SuchName$, ind% - 1)
    End If
'End If

flxTaxMuster(0).Rows = flxTaxMuster(0).FixedRows
If (SollTaxierTyp = MAG_SPEZIALITAET) Then
    sSQL = "Select GRP_HA.*, HA.Dichte FROM GRP_HA LEFT JOIN HA ON GRP_HA.PZN_2=HA.PZN WHERE GRP_HA.PZN_2=" + SollPzn
    On Error Resume Next
    ABDA_Komplett_Rec.Close
    Err.Clear
    On Error GoTo DefErr
    ABDA_Komplett_Rec.Open sSQL, ABDA_Komplett_Conn
    Do
        If (ABDA_Komplett_Rec.EOF) Then
            Exit Do
        End If
        
        pzn = CheckNullLong(ABDA_Komplett_Rec!PZN_1)
        
        If (ArtikelDbOk) Then
            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
'                    SQLStr = SQLStr + " AND LagerKz<>0"
            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
        Else
            FabsErrf% = ast.IndexSearch(0, pzn, FabsRecno&)
            If (FabsErrf% = 0) Then
                ast.GetRecord (FabsRecno& + 1)
            End If
        End If
        If (FabsErrf% = 0) Then
'                            ast.GetRecord (FabsRecno& + 1)
        Else
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
            If (TaxeRec.EOF = False) Then
                Call Taxe2ast(pzn$)
                FabsErrf% = 0
            End If
        End If
            
        If (FabsErrf% = 0) Then
            bOk = True
            If (SuchName <> "") Then
                bOk = (InStr(UCase(ast.KurzNeu), UCase(SuchName)) > 0)
            End If
            If (bOk) Then
                h = ast.pzn
                h = h + vbTab + Format(ast.aep, "0.00")
                h = h + vbTab + ast.MengNeu
                h = h + vbTab + ast.Meh
                h = h + vbTab + ast.KurzNeu
                If (SollTaxierTyp = MAG_SPEZIALITAET) Then
                    h = h + vbTab
                    h = h + vbTab + ast.herst
                    h = h + vbTab + ast.pzn
                End If
                flxTaxMuster(0).AddItem h, 1
            End If
        End If
        
        ABDA_Komplett_Rec.MoveNext
    Loop
    ABDA_Komplett_Rec.Close
End If

If (flxTaxMuster(0).Rows = 1) Then
    flxTaxMuster(0).AddItem vbTab + "keine passende PZN !"
    flxTaxMuster(1).Visible = False
    If (TaxMusterModus% = 0) Then cmdOk.Enabled = False
End If

Call DefErrPop
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Form_KeyDown")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'If (KeyCode = vbKeyF5) Then
'    cmdF5.Value = True
'End If
'
'Call DefErrPop
'End Sub

'Private Sub cmdF5_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdF5_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim erg%, ind%, row%
'Dim sXML$, h$
'
'If (cmdF5.Enabled = False) Then Call DefErrPop: Exit Sub
'
'row% = flxTaxMuster(0).row
'h$ = flxTaxMuster(0).TextMatrix(row%, 1)
'
'erg% = iMsgBox("Exportierte Rezeptur '" + h$ + "' löschen ?", vbYesNo Or vbInformation Or vbDefaultButton2)
'If (erg% = vbYes) Then
'    sXML = LennartzPfad + "\" + flxTaxMuster(0).TextMatrix(row, 0)
'    On Error Resume Next
'    Kill sXML
'    On Error GoTo DefErr
'
'    flxTaxMuster(1).Redraw = False
'
'    With flxTaxMuster(0)
'        .Redraw = False
'
'        .Rows = 1
'        Call flxTaxMusterBefuellen
'
'        .row = 1
'        .col = 1
'        .RowSel = .Rows - 1
'        .ColSel = .col
'        .Sort = 5
'        .col = 0
'        .ColSel = .Cols - 1
'        .Redraw = True
'        .SetFocus
'    End With
'
'    flxTaxMuster(1).Redraw = True
'End If
'
'Call DefErrPop
'End Sub

'Private Sub ZeigeTaxierZeile(row%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ZeigeTaxierZeile")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim iFlag%
'Dim h$, h2$
'
'With TaxierRec
'    iFlag% = .flag
'    If (iFlag% >= MAG_NN) Then
'        iFlag% = iFlag% - MAG_NN
'    End If
'
'    flxTaxMuster(1).TextMatrix(row%, 0) = Format(.ActPreis, "0.00")
'
'    h$ = Format(.ActMenge, "0.000")
'    Do
'        If (Right$(h$, 1) = ",") Then
'            h$ = Left$(h$, Len(h$) - 1)
'            Exit Do
'        End If
'        If (Right$(h$, 1) = "0") Then
'            h$ = Left$(h$, Len(h$) - 1)
'        Else
'            Exit Do
'        End If
'    Loop
'    flxTaxMuster(1).TextMatrix(row%, 1) = h$
'
'    flxTaxMuster(1).TextMatrix(row%, 2) = .Meh
'
'    h2$ = iTrim$(.kurz)
''    Call OemToChar(h2$, h2$)
'    flxTaxMuster(1).TextMatrix(row%, 3) = h2$
'
'    h2$ = ""
'    If (Val(.pzn) > 0) Then h2$ = .pzn
'    flxTaxMuster(1).TextMatrix(row%, 4) = h2$
'
''    flxTaxMuster(1).TextMatrix(row%, 5) = ""
''    flxTaxMuster(1).TextMatrix(row%, 6) = Format(iFlag%, "0")
''    flxTaxMuster(1).TextMatrix(row%, 7) = Format(.Kp, "0.00")
''    flxTaxMuster(1).TextMatrix(row%, 8) = Format(.Gstufe, "0.00")
'End With
'
'With flxTaxMuster(1)
'    .FillStyle = flexFillRepeat
'    .row = row%
'    .col = 0
'    .RowSel = .row
'    .ColSel = .Cols - 1
'    .CellForeColor = MagDarstellung&(iFlag%, 0)
'    .CellBackColor = MagDarstellung&(iFlag%, 1)
'
'    If (TaxierRec.flag >= MAG_NN) Then
'        .CellFontUnderline = True
'    Else
'        .CellFontUnderline = False
'    End If
'    .FillStyle = flexFillSingle
'End With
'
'Call DefErrPop
'End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseDown")
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
    
If (y <= wpara.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
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

Call DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Resize")
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

If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If

Call DefErrPop
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

'Private Sub nlcmdf5_Click()
'Call cmdF5_Click
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
    End If
End If
    
If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
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

