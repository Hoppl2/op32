VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmMagDarstellung 
   AutoRedraw      =   -1  'True
   Caption         =   "Darstellung des Magistralen Taxierens"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5970
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4200
      Picture         =   "MagDarstellung.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3960
      Picture         =   "MagDarstellung.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3720
      Picture         =   "MagDarstellung.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdStandard 
      Caption         =   "&Standard"
      Height          =   450
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1200
   End
   Begin VB.CommandButton cmdHintergrund 
      Caption         =   "&Hintergrund ..."
      Height          =   450
      Left            =   3120
      TabIndex        =   2
      Top             =   1200
      Width           =   1200
   End
   Begin VB.CommandButton cmdVordergrund 
      Caption         =   "&Vordergrund ..."
      Height          =   450
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxDarstellung 
      Height          =   2700
      Left            =   360
      TabIndex        =   0
      Top             =   240
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
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdStandard 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdHintergrund 
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdVordergrund 
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmMagDarstellung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "MAGDARSTELLUNG.FRM"

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

FormErg% = True
Call SpeicherIniDarstellung
Unload Me

Call DefErrPop
End Sub

Private Sub cmdVordergrund_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdVordergrund_Click")
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
Dim erg%

Call EditFarbe(0)
flxDarstellung.SetFocus

Call DefErrPop
End Sub

Private Sub cmdHintergrund_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdHintergrund_Click")
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
Dim erg%

Call EditFarbe(1)
flxDarstellung.SetFocus

Call DefErrPop
End Sub

Private Sub cmdStandard_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdStandard_Click")
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
Dim i%

With flxDarstellung
    .Redraw = False
    For i% = 0 To 6
        .row = i% + 1
        .col = 1
        .CellForeColor = .ForeColor
        .CellBackColor = .BackColor
    Next i%
    .Redraw = True
    .row = 1
End With
        
flxDarstellung.SetFocus

Call DefErrPop
End Sub

Private Sub flxDarstellung_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxDarstellung_GotFocus")
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

With flxDarstellung
    .col = 0
    .ColSel = .col
End With

Call DefErrPop
End Sub

Private Sub flxDarstellung_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxDarstellung_KeyPress")
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
    With flxDarstellung
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
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
            .ColSel = .col
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub flxDarstellung_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxDarstellung_RowColChange")
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

With flxDarstellung
    If (.Redraw) And (.Visible) Then
        .col = 0
        .ColSel = .col
    End If
End With

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

Call wpara.InitFont(Me)

With flxDarstellung
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Typ|<Darstellung"
    .Rows = 1
    
    .AddItem "Gefäß" + vbTab + "PIPETTENGLAS ECKIG"
    .AddItem "Arbeit" + vbTab + "LÖSEN K,MISCH FLÜSS,TEEMISCH"
    .AddItem "Spezialität" + vbTab + "KERASAL BASISSALBE"
    .AddItem "Sonstiges" + vbTab + "QUALITÄTSZUSCHLAG"
    .AddItem "Hilfstaxe" + vbTab + "DEXAMETHASONUM MIKROFEIN"
    .AddItem "Preis" + vbTab + "PREIS-EINGABE"
    .AddItem "Fam=Substanz" + vbTab + "KERASAL BASISSALBE"
    
    .ColWidth(0) = TextWidth(String(12, "A"))
    .ColWidth(1) = TextWidth(String(28, "A"))
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    
    For i% = 1 To (.Rows - 1)
        .row = i%
        .col = 1
        .CellForeColor = MagDarstellung&(i% - 1, 0)
        .CellBackColor = MagDarstellung&(i% - 1, 1)
    Next i%
    .row = 1
End With

Font.Bold = False   ' True

With cmdVordergrund
    .Left = flxDarstellung.Left + flxDarstellung.Width + 300
    .Top = flxDarstellung.Top
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
End With

With cmdHintergrund
    .Left = cmdVordergrund.Left
    .Top = cmdVordergrund.Top + cmdVordergrund.Height + 150
    .Width = cmdVordergrund.Width
    .Height = cmdVordergrund.Height
End With

With cmdStandard
    .Left = cmdVordergrund.Left
    .Top = cmdHintergrund.Top + cmdHintergrund.Height + 450
    .Width = cmdVordergrund.Width
    .Height = cmdVordergrund.Height
End With


Me.Width = cmdVordergrund.Left + cmdVordergrund.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height

cmdOk.Top = flxDarstellung.Top + flxDarstellung.Height + 150
cmdEsc.Top = cmdOk.Top

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxDarstellung
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = vbWhite
        .BackColorFixed = RGB(199, 176, 123)
        If (.SelectionMode = flexSelectionFree) Then
            .BackColorSel = RGB(135, 61, 52)
            .ForeColorSel = vbWhite '.ForeColor
        Else
            .BackColorSel = RGB(232, 217, 172)
            .ForeColorSel = .ForeColor
        End If
        .Appearance = 0
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    End With
    
    cmdVordergrund.Left = cmdVordergrund.Left + 2 * iAdd
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdOk.Top
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    Height = Height + iAdd2
    Width = Width + iAdd2 + 600
    
    With nlcmdOk
        .Init
'        .Left = (Me.ScaleWidth - 2 * .Width - 300)
'        .Top = tabProfil.Top + tabProfil.Height + iAdd + 600
        .Top = flxDarstellung.Top + flxDarstellung.Height + iAdd + 600
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
    
    With nlcmdVordergrund
        .Init
        .AutoSize = True
        .Left = cmdVordergrund.Left
        .Top = cmdVordergrund.Top
        .Caption = cmdVordergrund.Caption
        .TabIndex = cmdVordergrund.TabIndex
        .Enabled = cmdVordergrund.Enabled
        .Default = cmdVordergrund.Default
        .Cancel = cmdVordergrund.Cancel
        .Visible = True
    End With
    cmdVordergrund.Visible = False
    
    With nlcmdHintergrund
        .Init
        .Left = nlcmdVordergrund.Left
        .Top = nlcmdVordergrund.Top + nlcmdVordergrund.Height + 90
        .Caption = cmdHintergrund.Caption
        .TabIndex = cmdHintergrund.TabIndex
        .Enabled = cmdHintergrund.Enabled
        .Default = cmdHintergrund.Default
        .Cancel = cmdHintergrund.Cancel
        .Visible = True
        .Width = nlcmdVordergrund.Width
    End With
    cmdHintergrund.Visible = False
    
    With nlcmdStandard
        .Init
        .Left = nlcmdVordergrund.Left
        .Top = nlcmdHintergrund.Top + nlcmdHintergrund.Height + 90
        .Caption = cmdStandard.Caption
        .TabIndex = cmdStandard.TabIndex
        .Enabled = cmdStandard.Enabled
        .Default = cmdStandard.Default
        .Cancel = cmdStandard.Cancel
        .Visible = True
        .Width = nlcmdVordergrund.Width
    End With
    cmdStandard.Visible = False
    
    
    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxDarstellung
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

    On Error Resume Next
    For Each c In Controls
'        If (c.Container Is Me) Then
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
    
    '                If (c.Visible) Then
                        With c.Container
                            .ForeColor = RGB(180, 180, 180) ' vbWhite
                            .FillStyle = vbSolid
                            .FillColor = c.BackColor
    
                            RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                        End With
    '                End If
'                ElseIf (TypeOf c Is CheckBox) Then
'                    c.Height = 0
'                    c.Width = c.Height
'                    If (c.Name = "chkOptionen") Then
'                        If (c.index > 0) Then
'                            Load lblchkOptionen(c.index)
'                        End If
'                        With lblchkOptionen(c.index)
'                            .BackStyle = 0 'duchsichtig
'                            .Caption = c.Caption
'                            .Left = c.Left + 300
'                            .Top = c.Top
'                            .Width = TextWidth(.Caption) + 90
'                            .TabIndex = c.TabIndex
'                            .Visible = True
'                        End With
'                    ElseIf (c.Name = "chkOptionen2") Then
'                        If (c.index > 0) Then
'                            Load lblchkOptionen2(c.index)
'                        End If
'                        With lblchkOptionen2(c.index)
'                            .BackStyle = 0 'duchsichtig
'                            .Caption = c.Caption
'                            .Left = c.Left + 300
'                            .Top = c.Top
'                            .Width = TextWidth(.Caption) + 90
'                            .TabIndex = c.TabIndex
'                            .Visible = True
'                        End With
'                    End If
    '            ElseIf (TypeOf c Is OptionButton) Then
    '                c.Height = 0
    '                c.Width = c.Height
    '                If (c.Name = "optAbglFrage") Then
    '                    If (c.index > 0) Then
    '                        Load lblOptAbglFrage(c.index)
    '                    End If
    '                    With lblOptAbglFrage(c.index)
    '                        .BackStyle = 0 'duchsichtig
    '                        .Caption = c.Caption
    '                        .Left = c.Left + 300
    '                        .Top = c.Top
    '                        .Width = TextWidth(.Caption) + 90
    '                        .TabIndex = c.TabIndex
    '                        .Visible = True
    '                    End With
    '                End If
                End If
'            End If
        End If
    Next
    On Error GoTo DefErr
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdVordergrund.Visible = False
    nlcmdHintergrund.Visible = False
    nlcmdStandard.Visible = False
End If


Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Sub EditFarbe(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditFarbe")
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
Dim ind%
Dim l&, lColor&

On Error Resume Next

With flxDarstellung
    .Redraw = False
    .col = 1
    If (index = 0) Then
        lColor& = .CellForeColor
    Else
        lColor& = .CellBackColor
    End If
End With


With frmAction.dlg
    .color = lColor&
    .CancelError = True
    .Flags = cdlCCFullOpen + cdlCCRGBInit
    Call .ShowColor
    If (Err = 0) Then
        If (index = 0) Then
            flxDarstellung.CellForeColor = .color
        Else
            flxDarstellung.CellBackColor = .color
        End If
    End If
End With

With flxDarstellung
    .Redraw = True
    .col = 0
End With

Call DefErrPop
End Sub

Sub SpeicherIniDarstellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniDarstellung")
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
Dim i%
Dim l&
Dim h$, Key$

With flxDarstellung
    .Redraw = False
    For i% = 0 To 6
        .row = i% + 1
        .col = 1
        
'        MagDarstellung&(i%, 0) = 0
'        MagDarstellung&(i%, 1) = 0
'
'        If (.CellForeColor <> .ForeColor) Then MagDarstellung&(i%, 0) = .CellForeColor
'        If (.CellBackColor <> .BackColor) Then MagDarstellung&(i%, 1) = .CellBackColor
'
'        h$ = Hex$(MagDarstellung&(i%, 0)) + "," + Hex$(MagDarstellung&(i%, 1))
'        Key$ = "Darstellung" + Format(i%, "0")
'        l& = WritePrivateProfileString("Taxierung", Key$, h$, INI_DATEI)
        
        
        MagDarstellung&(i%, 0) = .CellForeColor
        MagDarstellung&(i%, 1) = .CellBackColor
        
        h$ = ""
        If (.CellForeColor <> .ForeColor) Then h$ = Hex$(MagDarstellung&(i%, 0))
        If (.CellBackColor <> .BackColor) Then h$ = h$ + "," + Hex$(MagDarstellung&(i%, 1))
        
        Key$ = "Darstellung" + Format(i%, "0")
        l& = WritePrivateProfileString("Taxierung", Key$, h$, INI_DATEI)
    Next i%
End With
       
Call DefErrPop
End Sub

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

Private Sub nlcmdVordergrund_Click()
Call cmdVordergrund_Click
End Sub

Private Sub nlcmdHintergrund_Click()
Call cmdHintergrund_Click
End Sub

Private Sub nlcmdStandard_Click()
Call cmdStandard_Click
End Sub

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



