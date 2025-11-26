VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmKkMatchcode 
   AutoRedraw      =   -1  'True
   Caption         =   "Matchcode - Auswahl"
   ClientHeight    =   5715
   ClientLeft      =   285
   ClientTop       =   315
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5880
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3960
      Picture         =   "KkMatchCode.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   4200
      Picture         =   "KkMatchCode.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4440
      Picture         =   "KkMatchCode.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtKKName 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxKassen 
      Height          =   2700
      Left            =   240
      TabIndex        =   2
      Top             =   720
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
   Begin VB.TextBox txtFlexBack 
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Tag             =   "0"
      Text            =   "Text1"
      Top             =   1440
      Width           =   1935
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1800
      TabIndex        =   5
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label lblKKName 
      Caption         =   "&Name/IK:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmKkMatchcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ANZ_ANZEIGE% = 15

Private Type KassenAnzeigeStruct
    Name As String * 100 '80
    DSnummer As Long
    Verweis As Long
    LagerKz As Byte
    pzn As Long
    Bookmark As Variant
    Key As String * 38  '10
    ZusatzTextKz As Byte
    BesorgtKz As Byte
End Type
Dim Ausgabe(ANZ_ANZEIGE%) As KassenAnzeigeStruct

'Dim KkassenRec As Recordset

Dim AnzAnzeige%

Dim iEditModus%

Private Const DefErrModul = "KKMATCHCODE.FRM"

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
Dim erg%, row%
Dim h$

If (ActiveControl.Name = txtKKName.Name) Then
    h$ = RTrim(UCase(txtKKName.text))
'    If (Left$(h$, 1) = "#") Then h$ = Mid$(h$, 2)
    erg% = SuchArtikel%(h$)
ElseIf (ActiveControl.Name = txtFlexBack.Name) Then
    With flxKassen
        AutIdemIk& = Val(.TextMatrix(.row, 0))
        AutIdemKbvNr& = Val(.TextMatrix(.row, 2))
    End With
    If (AutIdemIk& > 0) Then
        h$ = Format(AutIdemIk&, String(9, "0"))
        Call ActProgram.MakeActKkasse(h$)
        
        SQLStr = "SELECT count(*) AS iAnz FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
        SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
        SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            If (CheckNullLong(VdbBedRec!iAnz) = 1) Then
                SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
                SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
                SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
                FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
                If Not (VdbBedRec.EOF) Then
                    ActVebNr = CheckNullLong(VdbBedRec!VebNr)
                End If
            End If
        End If
        
        If (ActKasse% < 0) And (AutIdemIk& > 0) Then
            SQLStr = "SELECT Top 1* FROM VdbVIK"
            SQLStr = SQLStr + " WHERE IK=" + CStr(AutIdemIk)
            SQLStr = SQLStr + " ORDER BY KostentrNr"
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                ActKasse = CheckNullLong(VdbBedRec!KostentrNr)
            End If
        End If
    End If
    FormErg% = True
    Unload Me
End If

Call DefErrPop
End Sub

Private Sub flxKassen_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_DblClick")
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

Private Sub flxKassen_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKassen_GotFocus")
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

txtFlexBack.SetFocus

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
Dim i%, erg%, Breite1%
Dim iAdd%, iAdd2%
Dim l&
Dim h$
Dim c As Control

iEditModus = 1

FormErg% = False

Caption = "Auswahl Krankenkasse"

Me.KeyPreview = True
Call wpara.InitFont(Me)

With lblKKName
    .Top = 2 * wpara.TitelY
    .Left = wpara.LinksX
End With
With txtKKName
    .Left = lblKKName.Left + lblKKName.Width + 150
    .Top = lblKKName.Top
    .Width = TextWidth(String(20, "X"))
End With

With flxKassen
    .Rows = 2
    .FixedRows = 1
    .FormatString = ">IK|<Name|"
    .Rows = 1
    .ColWidth(0) = TextWidth(String(9, "9"))
    .ColWidth(1) = TextWidth(String(40, "A"))
    .ColWidth(2) = 0
'    .ColWidth(3) = wpara.FrmScrollHeight
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    
    .Height = .RowHeight(0) * (ANZ_ANZEIGE% + 1) + 90
    
    .Left = wpara.LinksX
    
    .Top = lblKKName.Top + lblKKName.Height + 300
End With

Font.Bold = False   ' True

Me.Width = flxKassen.Left + flxKassen.Width + 2 * wpara.LinksX

With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.Width - (.Width * 2 + 300)) / 2
    .Top = flxKassen.Top + flxKassen.Height + 300
End With
With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = cmdOk.Left + .Width + 300
    .Top = cmdOk.Top
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Font.Name = wpara.FontName(0)
Me.Font.Size = wpara.FontSize(0)

Call ErstBefüllung

With flxKassen
    .col = 0
    .ColSel = .Cols - 1
End With

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxKassen
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
    
        txtFlexBack.Top = .Top + 300
    End With
    
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
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
    
    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxKassen.Top + flxKassen.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdOk
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdEsc.Top
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxKassen
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

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

'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor

                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
'            ElseIf (TypeOf c Is CheckBox) Then
'                c.Height = 0
'                c.Width = c.Height
'                If (c.Name = "chkHistorie") Then
'                    If (c.Index > 0) Then
'                        Load lblchkHistorie(c.Index)
'                    End If
'                    With lblchkHistorie(c.Index)
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
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If
'''''''''
Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Sub MachMaske()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachMaske")
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
Dim i%, j%, erg%, h$, ind%, ind2%, sAnz%, NurLagerndeAktiv%
Dim OrgControl As Control

If (txtKKName.text = "") Then
    Call DefErrPop: Exit Sub
End If

Ausgabe(0).Name = Left$(Ausgabe(0).Name, Ausgabe(0).Verweis)

AnzAnzeige% = ANZ_ANZEIGE%
For i% = 1 To ANZ_ANZEIGE%
    erg% = SuchWeiter%(i% - 1, True)
    If (erg%) Then
        Call Umspeichern(buf$, i%)
        Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis)
    Else
        AnzAnzeige% = i%
        Exit For
    End If
Next i%

Set OrgControl = ActiveControl
Call AuswahlBefüllen
OrgControl.SetFocus
With flxKassen
    .row = 1
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub

Private Sub txtFlexBack_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtFlexBack_GotFocus")
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

With flxKassen
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub txtFlexBack_lostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtFlexBack_lostFocus")
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

With flxKassen
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub txtFlexBack_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtFlexBack_KeyDown")
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
Dim erg%, col%, iLaenge%

Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
        Call AuswahlRowChange(KeyCode)
        KeyCode = 0
    Case vbKeyLeft
        KeyCode = 0
    Case vbKeyRight
        KeyCode = 0
End Select

With flxKassen
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub

Private Sub txtFlexBack_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtFlexBack_KeyPress")
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

Call DefErrPop
End Sub

Sub AuswahlRowChange(KeyCode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlRowChange")
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
Dim erg%, i%, j%, h$, ind%, neu%, NurLagerndeAktiv%, sAnz%
    
neu% = True
With flxKassen
    Select Case KeyCode
        Case vbKeyUp
            If (.row > 1) Then
                .row = .row - 1
                neu% = False
            Else
                erg% = SuchWeiter%(0, False)
                If (erg%) Then
                    GoSub ZeileHinein
                Else
                    .row = 1
                End If
            End If
        Case vbKeyPageUp
            erg% = SuchWeiter%(0, False)
            If (erg%) Then
                For j% = 1 To (ANZ_ANZEIGE - 1)
                    GoSub ZeileHinein
                    erg% = SuchWeiter%(0, False)
                    If (erg% = False) Then Exit For
                Next j%
            Else
                .row = 1
            End If
        Case vbKeyDown
            If (.row < .Rows - 1) Then
                .row = .row + 1
                neu% = False
            Else
                erg% = SuchWeiter%(AnzAnzeige% - 1, True)
                If (erg%) Then
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    i% = AnzAnzeige% - 1
                    Call Umspeichern(buf$, i%)
                    Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis)
                Else
                    .row = AnzAnzeige%
                End If
            End If
        Case vbKeyPageDown
            erg% = SuchWeiter%(AnzAnzeige% - 1, True)
            If (erg%) Then
                For j% = 1 To AnzAnzeige%
                    For i% = 1 To (AnzAnzeige% - 1)
                        Ausgabe(i% - 1) = Ausgabe(i%)
                    Next i%
                    i% = AnzAnzeige% - 1
                    Call Umspeichern(buf$, i%)
                    Ausgabe(i%).Name = Left$(Ausgabe(i%).Name, Ausgabe(i%).Verweis)
                    erg% = SuchWeiter%(AnzAnzeige% - 1, True)
                    If (erg% = False) Then Exit For
                Next j%
            Else
                .row = AnzAnzeige%
            End If
    End Select
End With


If (neu% = True) Then Call AuswahlBefüllen
txtFlexBack.SetFocus
Call DefErrPop: Exit Sub

ZeileHinein:
For i% = (ANZ_ANZEIGE - 2) To 1 Step -1
    Ausgabe(i%) = Ausgabe(i% - 1)
Next i%
Call Umspeichern(buf$, 0)
Ausgabe(0).Name = Left$(Ausgabe(0).Name, Ausgabe(0).Verweis)
If (AnzAnzeige% < (ANZ_ANZEIGE - 1)) Then
    AnzAnzeige% = AnzAnzeige% + 1
    flxKassen.Rows = AnzAnzeige% + 1
End If
Return

Call DefErrPop
End Sub

Private Sub txtKKName_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtKKName_GotFocus")
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

With txtKKName
    .SelStart = 0
    .SelLength = Len(.text)
End With

iEditModus = 1

Call DefErrPop
End Sub

Function SuchArtikel%(Such$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SuchArtikel%")
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
Dim i%, l%, l2%, ind%, TextSuche%, IstPzn%, LeftLen%, InhMengeFlag%, gef%, IstCode%
Dim InhMenge&
Dim SQLStr$, nSuch$, ch$, h$, oMatchModus%
Dim iBookmark As Variant

SuchArtikel% = True

If (Such$ = "") Then Such$ = Space$(10)
l% = Len(Such$)

IstPzn% = True
For i% = 1 To l%
    ch$ = Mid$(Such$, i%, 1)
    If (InStr("0123456789", ch$) = 0) Then
        IstPzn% = False
        Exit For
    End If
Next i%

If (IstPzn% = True) Then
    If (Len(Such) = 9) Then
        If (Left(Such, 2) = "10") Then
            Such = Mid(Such, 3)
        End If
    End If
    SQLStr$ = "SELECT * FROM KKassen WHERE Ik = " + Such$
'    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'    If (KkassenRec.RecordCount > 0) Then
    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
    If (KkassenRec.EOF) Then
        AutIdemIk& = Val(Such)
        AutIdemKbvNr& = 0
    Else
        AutIdemIk& = CheckNullLong(KkassenRec!IK)
        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
    End If
    If (AutIdemIk& > 0) Then
        h$ = Format(AutIdemIk&, String(9, "0"))
        Call ActProgram.MakeActKkasse(h$)
        
        SQLStr = "SELECT count(*) AS iAnz FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
        SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
        SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            If (CheckNullLong(VdbBedRec!iAnz) = 1) Then
                SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
                SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
                SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
                FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
                If Not (VdbBedRec.EOF) Then
                    ActVebNr = CheckNullLong(VdbBedRec!VebNr)
                End If
            End If
        End If
        
        If (ActKasse% < 0) And (AutIdemIk& > 0) Then
            SQLStr = "SELECT Top 1* FROM VdbVIK"
            SQLStr = SQLStr + " WHERE IK=" + CStr(AutIdemIk)
            SQLStr = SQLStr + " ORDER BY KostentrNr"
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                ActKasse = CheckNullLong(VdbBedRec!KostentrNr)
            End If
        End If
        
        FormErg% = True
        Unload Me
        Call DefErrPop: Exit Function
    End If
End If

'Set KkassenRec = kKassenDB.OpenRecordset("Kkassen", dbOpenTable)
'KkassenRec.index = "Name"
'KkassenRec.Seek ">=", Such$
'If (KkassenRec.NoMatch = True) Then
On Error Resume Next
KkassenRec.Close
On Error GoTo DefErr
SQLStr = "SELECT * FROM " + "kkassen" + " WHERE Name LIKE '" + Such + "%'"
SQLStr = SQLStr + " ORDER BY Name"
KkassenRec.Open SQLStr$, kKassenDB.ActiveConn
If (KkassenRec.EOF) Then
    SuchArtikel% = False
Else
'        Call Umspeichern("", 0)
End If

If (SuchArtikel% = True) Then
    Call SuchArtikelTrue(IstPzn%)
End If

Call DefErrPop
End Function

Sub SuchArtikelTrue(PznFlag%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SuchArtikelTrue")
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
Dim j%, ind%
Dim h$

Call Umspeichern("", 0)
'Call MachAuswahlGrid
Call MachMaske
With flxKassen
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .SetFocus
End With

Call DefErrPop
End Sub

Sub AuswahlBefüllen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlBefüllen")
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
Dim i%, j%, k%, ind%, AltRow%, AltCol%, AnzArbeitCols%, OrgCols%, iKz%, aFontBold%
Dim Suc&
Dim h$

'On Error Resume Next

With flxKassen
    txtFlexBack.Visible = False
     .Visible = False
    AltRow% = .row
    AltCol% = .col
    
    .Rows = AnzAnzeige% + 1
    
    For i% = 1 To AnzAnzeige%
        h$ = Ausgabe(i% - 1).Name
        For j% = 0 To .Cols - 2
            ind% = InStr(h$, vbTab)
            .TextMatrix(i%, j%) = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Next j%
        .TextMatrix(i%, .Cols - 1) = RTrim$(h$)
    Next i%
        
    .Rows = AnzAnzeige% + 1
    
    .row = AltRow%
    .col = AltCol%
    .Visible = True
    txtFlexBack.Visible = True
End With

Call DefErrPop
End Sub

Function SuchWeiter%(ind%, Typ%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SuchWeiter%")
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
Dim AltRecno&
Dim h$, AltKey$

SuchWeiter% = False

KkassenRec.Bookmark = Ausgabe(ind%).Bookmark
If (Typ%) Then
    KkassenRec.MoveNext
    If (KkassenRec.EOF = False) Then SuchWeiter% = True
Else
    KkassenRec.MovePrevious
    If (KkassenRec.BOF = False) Then SuchWeiter% = True
End If

Call DefErrPop
End Function

Sub Umspeichern(buf$, pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Umspeichern")
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
Dim ind%, ind2%, i%, z%, z1%, z2%, dAnzFields%, IstMerkzettelMdb%
Dim zu!
Dim EK#, vk#
Dim h$, h2$, char$, zuza$, zusa$, SQLStr$

Ausgabe(pos%).LagerKz = 1
Ausgabe(pos%).ZusatzTextKz = 1
Ausgabe(pos%).BesorgtKz = 1

Ausgabe(pos%).pzn = KkassenRec!IK

Ausgabe(pos%).Name = Format(KkassenRec!IK, "0") + vbTab + KkassenRec!Name + vbTab + Format(KkassenRec!KbvNr, "0") + vbTab
Ausgabe(pos%).Verweis = Len(RTrim(Ausgabe(pos%).Name))

Ausgabe(pos%).LagerKz = 0
Ausgabe(pos%).Bookmark = KkassenRec.Bookmark
    
Call DefErrPop
End Sub

Private Sub ErstBefüllung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErstBefüllung")
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
Dim h$, SQLStr$
Dim OrgControl As Control

SQLStr$ = "SELECT * FROM AutIdemKrankenkassen ORDER BY AnzRezepte DESC"
Set AutIdemKkRec = RezSpeicherDB.OpenRecordset(SQLStr$)
AnzAnzeige% = 0
'For i% = 1 To 10
'    If (AutIdemKkRec.EOF) Then
'        Exit For
'    End If
'
'    SQLStr$ = "SELECT * FROM Kkassen WHERE Ik=" + Format(AutIdemKkRec!IK, "0")
''    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
''    If (KkassenRec.RecordCount > 0) Then
'    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
'    If Not (KkassenRec.EOF) Then
'        Call Umspeichern("", i% - 1)
'        AnzAnzeige% = i%
'    End If
'
'    AutIdemKkRec.MoveNext
'Next i%
Do
    If (AutIdemKkRec.EOF) Then
        Exit Do
    End If
    
    SQLStr$ = "SELECT * FROM Kkassen WHERE Ik=" + Format(AutIdemKkRec!IK, "0")
'    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'    If (KkassenRec.RecordCount > 0) Then
    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
    If Not (KkassenRec.EOF) Then
        Call Umspeichern("", AnzAnzeige%)
        AnzAnzeige% = AnzAnzeige% + 1
        If (AnzAnzeige >= 10) Then
            Exit Do
        End If
    End If
    
    AutIdemKkRec.MoveNext
Loop
If (AnzAnzeige% > 0) Then
    'Set OrgControl = ActiveControl
    Call AuswahlBefüllen
    'OrgControl.SetFocus
    With flxKassen
        .FillStyle = flexFillRepeat
        .row = 1
        .col = 1
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 5
        .FillStyle = flexFillSingle
        
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
    End With
End If

If (AutIdemIk& > 0) Then
    txtKKName = Format(AutIdemIk&, "0")
End If

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

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
'    ElseIf (KeyAscii = Asc("<")) And (nlcmdImport(0).Visible) Then
''        Call nlcmdChange_Click(0)
'        nlcmdImport(0).Value = 1
'    ElseIf (KeyAscii = Asc(">")) And (nlcmdImport(1).Visible) Then
''        Call nlcmdChange_Click(1)
'        nlcmdImport(1).Value = 1
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





